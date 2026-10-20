#!/usr/bin/env node

// Offline checks against the production functions, exposed only in this VM.
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const source = fs.readFileSync(path.join(__dirname, '../dist/ui.js'), 'utf8');
const marker = '  startInitialization();\n})();';
assert(source.includes(marker));
const element = () => {
  const listeners = new Map();
  return {
    children: [], options: [], dataset: {}, classList: { toggle() {} },
    addEventListener(type, handler) {
      if (!listeners.has(type)) listeners.set(type, new Set());
      listeners.get(type).add(handler);
    },
    removeEventListener(type, handler) { listeners.get(type)?.delete(handler); },
    fire(type) { for (const handler of Array.from(listeners.get(type) || [])) handler({ preventDefault() {} }); },
    appendChild(child) { this.children.push(child); },
    replaceChildren(...children) { this.children = children.flatMap((child) => child.fragment ? child.children : [child]); },
    focus() { document.activeElement = this; },
    querySelector() { return null; },
    showModal() { this.open = true; },
    close() { this.open = false; this.fire('close'); }
  };
};
const elements = new Map();
const document = {
  referrer: '', activeElement: null,
  getElementById(id) {
    if (!elements.has(id)) elements.set(id, element());
    return elements.get(id);
  },
  createElement: element,
  createDocumentFragment: () => ({ ...element(), fragment: true })
};
const window = { location: { origin: 'https://ado.example', search: '' }, addEventListener() {} };
const context = { window, document, URL, URLSearchParams, setTimeout, clearTimeout, console };
vm.runInNewContext(source.replace(marker,
  '  window.hooks = { buildYamlDiff, confirmPipelineOverwrite, prepareGeneratedFiles, postScaffold, postGeneratedFiles, handleSubmit, state, buildPipelineFilename };\n})();'
), context);
const { buildYamlDiff, confirmPipelineOverwrite, prepareGeneratedFiles, postScaffold, postGeneratedFiles, handleSubmit, state, buildPipelineFilename } = window.hooks;

const endings = { LF: '\n', CRLF: '\r\n', CR: '\r', EOF: '' };
const reconstruct = (rows, excluded) => rows.filter(({ kind }) => kind !== excluded)
  .map(({ text, ending }) => text + endings[ending]).join('');
for (const [before, after] of [
  ['', ''], ['', 'trigger: none\n'], ['trigger: none\n', ''],
  ['a\nb\nc\n', 'a\nnew\nc\n'], ['same\nsame\nend', 'same\nend\nsame'],
  ['key: value\n', 'key: value'], ['  script: x\r\n', '\tscript: x\n'],
  ['a\rb\r', 'a\nb\n'], ['<script>alert(1)</script>\n', '${VALUE}\n'],
  [Array.from({ length: 1100 }, (_, i) => `old ${i}\n`).join(''), Array.from({ length: 1100 }, (_, i) => `new ${i}\n`).join('')]
]) {
  const rows = buildYamlDiff(before, after);
  assert.strictEqual(reconstruct(rows, 'added'), before);
  assert.strictEqual(reconstruct(rows, 'removed'), after);
  assert.deepStrictEqual(Array.from(rows.filter(({ oldLine }) => oldLine !== null), ({ oldLine }) => oldLine),
    Array.from({ length: rows.filter(({ oldLine }) => oldLine !== null).length }, (_, i) => i + 1));
}
assert.deepStrictEqual(Array.from(buildYamlDiff('one\nold\nthree\n', 'one\nnew\nthree\n'), ({ kind }) => kind),
  ['context', 'removed', 'added', 'context']);

const hostUri = 'https://ado.example/Collection/';
const base = { hostUri, projectId: 'project', repoId: 'repo', repositoryName: 'Project_Azure_DevOps', accessToken: 'synthetic-test-value' };
const commit = '1'.repeat(40);
let files;
let writes;
let conflict;
const resetServer = (initialFiles) => {
  files = initialFiles;
  writes = [];
  conflict = false;
  context.fetch = async (url, options = {}) => {
    let status = 200;
    let body;
    const parsed = new URL(url);
    if (url.includes('/refs?')) body = { value: [{ objectId: commit }] };
    else if (url.includes('/items?')) {
      assert.strictEqual(parsed.searchParams.get('versionDescriptor.version'), commit);
      assert.strictEqual(parsed.searchParams.get('versionDescriptor.versionType'), 'commit');
      body = files[parsed.searchParams.get('path')];
      if (body === undefined) { status = 404; body = ''; }
    } else if (url.includes('/pushes?') && options.method === 'POST') {
      const push = JSON.parse(options.body);
      assert.strictEqual(push.refUpdates[0].oldObjectId, commit);
      writes.push(push);
      status = conflict ? 409 : 200;
      body = conflict ? { message: 'Branch changed during review' } : { pushId: 1 };
    } else throw new Error(`Unexpected request: ${url}`);
    return {
      status, ok: status < 400, headers: { get() {} },
      async json() { return body; }, async text() { return typeof body === 'string' ? body : JSON.stringify(body); }
    };
  };
};
const normal = (confirmOverwrite) => postScaffold({ ...base, pipelineFilename: 'pipeline.yml', content: 'new\n', ...(confirmOverwrite ? { confirmOverwrite } : {}) });
const mono = (confirmOverwrite) => postGeneratedFiles({ ...base, files: [
  { path: '/pipeline.yml', content: 'new\n' },
  { path: '/.devops/deployments.yml', content: 'starter\n', overwrite: false }
], ...(confirmOverwrite ? { confirmOverwrite } : {}) });
const tick = () => new Promise((resolve) => setImmediate(resolve));

(async () => {
  for (const save of [normal, mono]) {
    for (const initial of [{}, { '/pipeline.yml': 'new\n', '/.devops/deployments.yml': 'operator\n' }]) {
      resetServer(initial);
      await save(async () => { throw new Error('New/unchanged YAML must not ask for overwrite approval'); });
      assert.strictEqual(writes.length, initial['/pipeline.yml'] ? 0 : 1);
    }
    resetServer({ '/pipeline.yml': 'old\n' });
    let approve;
    let reviewed;
    const pending = save((details) => {
      reviewed = details;
      return new Promise((resolve) => { approve = resolve; });
    });
    await tick();
    assert.strictEqual(writes.length, 0, 'No push while approval is pending');
    assert.strictEqual(reviewed.previous, 'old\n');
    assert.strictEqual(reviewed.proposed, 'new\n');
    assert.strictEqual(reviewed.filePath, '/pipeline.yml');
    assert.strictEqual(reviewed.repositoryName, base.repositoryName);
    assert.strictEqual(reviewed.branchName, 'main');
    approve(true);
    await pending;
    assert.strictEqual(writes.length, 1);
    assert.strictEqual(writes[0].commits[0].changes[0].newContent.content, reviewed.proposed);
    for (const answer of [false, undefined, 'true']) {
      resetServer({ '/pipeline.yml': 'old\n' });
      await assert.rejects(save(async () => answer), (error) => error.code === 'PIPELINE_OVERWRITE_CANCELLED');
      assert.strictEqual(writes.length, 0, 'Cancel aborts the entire commit, including MR contract additions');
    }
    resetServer({ '/pipeline.yml': 'old\n' });
    conflict = true;
    await assert.rejects(save(async () => true), (error) => error.status === 409);
    assert.strictEqual(writes.length, 1, 'Do not retry a concurrent write with stale approval');
    files['/pipeline.yml'] = 'changed after first review\n';
    conflict = false;
    await assert.rejects(save(async ({ previous }) => {
      assert.strictEqual(previous, 'changed after first review\n');
      return false;
    }), (error) => error.code === 'PIPELINE_OVERWRITE_CANCELLED');
    assert.strictEqual(writes.length, 1, 'A retry requires a fresh review');
  }

  resetServer({ '/pipeline.yml': 'old\n' });
  let reviewCount = 0;
  const reviewedPlan = await prepareGeneratedFiles({ ...base, files: [{ path: '/pipeline.yml', content: 'new\n' }],
    confirmOverwrite: async () => { reviewCount += 1; return true; }
  });
  assert(Object.isFrozen(reviewedPlan));
  assert.strictEqual(writes.length, 0, 'Preparing and confirming a plan is read-only');
  await postScaffold({ ...base, pipelineFilename: 'pipeline.yml', content: 'new\n', reviewedPlan,
    confirmOverwrite: async () => { throw new Error('Do not prompt again after preflight'); }
  });
  assert.strictEqual(reviewCount, 1);
  assert.strictEqual(writes.length, 1);
  await assert.rejects(postScaffold({ ...base, pipelineFilename: 'pipeline.yml', content: 'different proposal\n', reviewedPlan }),
    (error) => error.code === 'PIPELINE_REVIEW_STALE');
  assert.strictEqual(writes.length, 1, 'Approval cannot transfer to different proposed content');

  // Drive the actual submit handler: cancellation must precede every resource
  // mutation, rather than merely protecting the eventual YAML push.
  const formPayload = { service: 'api', environment: 'pro', stack: 'default', pool: 'Default',
    dockerfileDir: '.', repositoryAddress: 'registry.example', containerRegistryService: 'Registry', komodoServer: 'Production' };
  context.FormData = class { entries() { return Object.entries(formPayload); } };
  window.PipelineGeneratorReleaseConfig = { enabled: false };
  const submitScenario = ({ mode, newRepository = false, appearAfterPreflight = false }) => {
    Object.assign(state, { ...base, mode, rawProjectName: 'Project', projectName: 'Project',
      rawRepositoryName: 'Project_Backend', repositoryName: 'Project_Backend', sourceBranch: 'main',
      accessToken: 'synthetic-test-value', isSubmitting: false, provisioningComplete: false,
      deploymentTargetsReady: true, deploymentTargets: { environmentConfigs: [{ name: 'pro', domain: 'example.com' }] }
    });
    const filename = buildPipelineFilename({ projectName: 'Project', repositoryName: 'Project_Backend',
      service: formPayload.service, environment: 'pro', branchName: 'main', komodoServer: 'Production', mode });
    const repos = new Map(newRepository ? [] : [['repo', { id: 'repo', name: base.repositoryName }]]);
    const scenario = { requests: [], mutations: [], head: newRepository ? null : commit, pushes: [], listings: 0 };
    context.fetch = async (url, options = {}) => {
      const method = options.method || 'GET';
      const request = { url, method, body: options.body ? JSON.parse(options.body) : null };
      scenario.requests.push(request);
      const parsed = new URL(url);
      let body;
      let status = 200;
      if (url.includes('path=%2Fkomodo-servers-creds.env')) {
        body = 'KOMODO_ADDRESS=https://komodo.example\nKOMODO_API_KEY=synthetic-read-key\nKOMODO_API_SECRET=synthetic-read-secret\n';
      } else if (url === 'https://komodo.example/read') {
        assert.strictEqual(request.body.type, 'ListDockerNetworks');
        body = [{ name: 'nginx-network' }];
      } else if (url.includes('/_apis/git/repositories?')) {
        if (method === 'GET') {
          scenario.listings += 1;
          if (appearAfterPreflight && scenario.listings > 1) {
            repos.set('repo', { id: 'repo', name: base.repositoryName });
            scenario.head = commit;
          }
          body = { value: Array.from(repos.values()) };
        } else {
          scenario.mutations.push(request);
          const name = request.body.name;
          const id = name === base.repositoryName ? 'repo' : name;
          body = { id, name };
          repos.set(id, body);
        }
      } else if (url.includes('/refs?')) {
        const head = url.includes('/repositories/repo/') ? scenario.head : null;
        body = { value: head ? [{ objectId: head }] : [] };
      } else if (url.includes('/items?')) {
        assert.strictEqual(parsed.searchParams.get('versionDescriptor.version'), commit);
        assert.strictEqual(parsed.searchParams.get('versionDescriptor.versionType'), 'commit');
        body = parsed.searchParams.get('path') === '/.devops/deployments.yml' ? 'operator contract\n' : 'old pipeline\n';
      } else if (url.includes('/pushes?')) {
        scenario.mutations.push(request);
        scenario.pushes.push(request.body);
        body = { pushId: scenario.pushes.length };
      } else if (method === 'PATCH' && url.includes('/_apis/git/repositories/')) {
        scenario.mutations.push(request);
        body = {};
      } else if (url.includes('/_apis/pipelines?')) {
        if (method === 'GET') body = { value: [] };
        else { scenario.mutations.push(request); body = { id: 21, name: filename }; }
      } else if (url.includes('/_apis/build/definitions?')) {
        body = { value: [] };
      } else throw new Error(`Unexpected full-flow request: ${method} ${url}`);
      return { status, ok: status < 400, headers: { get() {} },
        async json() { return body; }, async text() { return typeof body === 'string' ? body : JSON.stringify(body); } };
    };
    return scenario;
  };
  const dialog = document.getElementById('pipeline-overwrite-dialog');
  const cancel = document.getElementById('overwrite-cancel');
  const confirm = document.getElementById('overwrite-confirm');
  for (const mode of ['pipeline', 'monorepo']) {
    for (const action of ['cancel', 'escape', 'close', 'confirm', 'concurrent-change']) {
      const scenario = submitScenario({ mode });
      const pending = handleSubmit({ preventDefault() {} });
      await tick();
      assert.strictEqual(dialog.open, true, 'Submit must open the review before Step 1');
      assert(scenario.requests.every(({ method }) => method === 'GET'), 'Only read-only requests may precede consent');
      assert.strictEqual(scenario.mutations.length, 0, 'No repository/Compose/Nginx/Pipeline/Release changes before consent');
      const reviewedRows = document.getElementById('overwrite-diff-lines').children;
      const proposed = reviewedRows.filter((row) => row.className !== 'diff-removed')
        .map((row) => row.children[3].textContent + endings[row.children[4].textContent]).join('');
      if (action === 'cancel') cancel.fire('click');
      else if (action === 'escape') dialog.fire('cancel');
      else if (action === 'close') dialog.close();
      else {
        if (action === 'concurrent-change') scenario.head = '2'.repeat(40);
        confirm.fire('click');
      }
      await pending;
      if (action === 'confirm') {
        assert.strictEqual(state.provisioningComplete, true, document.getElementById('status').textContent);
        assert(scenario.mutations.length > 0);
        const yamlPush = scenario.pushes.find((push) => push.commits[0].changes.some(({ item }) => item.path.endsWith('.yml') && !item.path.includes('compose')));
        assert(yamlPush, 'Confirmed preflight continues through the actual YAML write');
        assert.strictEqual(yamlPush.refUpdates[0].oldObjectId, commit);
        assert.strictEqual(yamlPush.commits[0].changes[0].newContent.content, proposed);
        assert(!dialog.open, 'Step 2 must reuse the reviewed plan without a second prompt');
      } else {
        assert.strictEqual(scenario.mutations.length, 0, 'Cancellation/stale preflight must leave all resources untouched');
        assert.strictEqual(state.provisioningComplete, false);
        assert.strictEqual(state.isSubmitting, false);
      }
    }
    const appeared = submitScenario({ mode, newRepository: true, appearAfterPreflight: true });
    await handleSubmit({ preventDefault() {} });
    assert.strictEqual(appeared.mutations.length, 0, 'A concurrently created repository with existing YAML must stop before support writes');
    assert.strictEqual(state.provisioningComplete, false);
    const missing = submitScenario({ mode, newRepository: true });
    await handleSubmit({ preventDefault() {} });
    assert.strictEqual(state.provisioningComplete, true, document.getElementById('status').textContent);
    assert(missing.mutations.some(({ body }) => body.name === base.repositoryName), 'A missing repository is created after read-only preflight');
  }

  // Exercise the default dialog, not just an injected confirmation callback.
  const diffBody = document.getElementById('overwrite-diff-lines');
  for (const action of ['cancel', 'escape', 'close', 'confirm']) {
    resetServer({ '/pipeline.yml': '<img src=x onerror=alert(1)>\r\n' });
    const focused = element();
    focused.focus();
    const pending = normal();
    const rejected = action === 'confirm' ? pending : assert.rejects(pending, (error) => error.code === 'PIPELINE_OVERWRITE_CANCELLED');
    await tick();
    assert.strictEqual(dialog.open, true);
    assert.strictEqual(document.activeElement, cancel);
    assert.strictEqual(writes.length, 0);
    dialog.fire('close');
    assert.strictEqual(dialog.open, true, 'Ignore a queued close event from a previous review');
    assert(diffBody.children.some((row) => row.children.some((cell) => cell.textContent === '<img src=x onerror=alert(1)>')),
      'YAML must be rendered as text, never HTML');
    if (action === 'cancel') cancel.fire('click');
    if (action === 'escape') dialog.fire('cancel');
    if (action === 'close') dialog.close();
    if (action === 'confirm') confirm.fire('click');
    await rejected;
    assert.strictEqual(writes.length, action === 'confirm' ? 1 : 0);
    assert.strictEqual(dialog.open, false);
    assert.strictEqual(diffBody.children.length, 0, 'Clear reviewed source from the dialog on close');
    assert.strictEqual(document.activeElement, focused);
  }
  dialog.showModal = undefined;
  await assert.rejects(confirmPipelineOverwrite({ previous: 'old', proposed: 'new' }), /Cannot display/);
  console.log('Pipeline overwrite validation passed: read-only preflight before all mutations, normal/MR submit cancellation/approval, exact reviewed writes, stale-plan rejection, concurrent creation, and safe dialog rendering.');
})().catch((error) => { console.error(error); process.exitCode = 1; });
