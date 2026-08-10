#!/usr/bin/env node

/*
 * Focused behavioral regression tests for the browser provisioning logic.
 * The production IIFE is instrumented in-memory to expose selected functions;
 * no runtime test hook is shipped in dist/ui.js and no network call is made.
 */
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const uiPath = path.join(root, 'dist/ui.js');
const source = fs.readFileSync(uiPath, 'utf8');
const initializationMarker = '  startInitialization();\n})();';
if (!source.includes(initializationMarker)) {
  throw new Error('[ui behavior validation] Could not find the UI initialization marker.');
}

const instrumented = source.replace(
  initializationMarker,
  `  window.__PipelineGeneratorTestHooks = {
	    buildPipelineFilename,
	    buildPipelineName,
	    getAuthHeader,
	    getDialogConfiguration,
	    getHostNavigationState,
	    buildSignOutUrl,
	    buildExtensionManagementUrl,
	    normalizeAccessTokenError,
	    isHostAuthorizationError,
	    buildTokenRecoveryMessage,
	    openExtensionAuthorization,
	    restartAzureDevOpsSession,
	    resolveReleaseAgentQueue,
	    resolveReleaseVariableGroup,
	    resolveReleaseInlineScript,
	    postScaffold,
	    pipelineBindingMatches,
    upsertPipelineDefinition,
    ensureReleaseDefinition,
    state
  };
})();`
);

const element = (overrides = {}) => ({
  value: '',
  textContent: '',
  className: '',
  disabled: false,
  dataset: {},
  options: [],
  innerHTML: '',
  classList: {
    toggle() {}
  },
  addEventListener() {},
  focus() {},
  appendChild(child) {
    this.options.push(child);
  },
  ...overrides
});

const submitButton = element();
const form = element({
  querySelector: () => submitButton
});
const environment = element({
  value: 'demo',
  options: ['dev', 'demo', 'qa', 'pro'].map((value) => ({ value }))
});
const elements = new Map([
  ['branch-label', element()],
  ['branch', element()],
  ['environment', environment],
  ['pool', element()],
  ['service', element()],
  ['containerRegistryService', element()],
  ['dockerfileDir', element()],
  ['pipeline-form', form],
  ['status', element()],
  ['targetRepo', element()],
  ['komodoServer', element({ options: [] })],
  ['reauth-panel', element({ className: 'auth-fallback hidden' })],
  ['reauth-message', element()],
  ['authorize-extension', element()],
  ['reauthenticate', element()]
]);

const document = {
  referrer: '',
  getElementById: (id) => elements.get(id) || element(),
  createElement: () => element(),
  head: { appendChild() {} }
};
const window = {
  location: {
    origin: 'https://azure.example.local',
    href: 'https://azure.example.local/extension/dist/index.html',
    search: ''
  },
  opener: null,
  addEventListener() {},
  setTimeout,
  PipelineGeneratorReleaseConfig: {
    enabled: true,
    folder: '\\komodo',
    environmentName: 'komodo',
    nameSuffix: '_Release',
    bashTaskName: 'Run Komodo deployment',
    variableGroupName: 'KomodoAPI',
    requiredVariableNames: ['AZP_TOKEN', 'KOMODO_API_KEY', 'KOMODO_API_SECRET'],
    scriptSource: { type: 'inline', content: '#!/usr/bin/env bash\necho regression-test' }
  }
};
window.parent = window;

const quietConsole = {
  log() {},
  warn() {},
  error() {}
};
const context = {
  window,
  document,
  URL,
  URLSearchParams,
  FormData,
  setTimeout,
  clearTimeout,
  setInterval,
  clearInterval,
  console: quietConsole,
  fetch: async () => {
    throw new Error('Unexpected fetch call.');
  }
};
vm.runInNewContext(instrumented, context, { filename: 'dist/ui.js' });
const hooks = window.__PipelineGeneratorTestHooks;
assert(hooks, 'UI test hooks were not exposed by the in-memory instrumentation.');

const response = ({ status = 200, body = {}, url = 'https://azure.example.local/mock' }) => ({
  ok: status >= 200 && status < 300,
  status,
  url,
  headers: { get: () => '' },
  async json() {
    return body;
  },
  async text() {
    return typeof body === 'string' ? body : JSON.stringify(body);
  }
});

const hostUri = 'https://azure.example.local/DefaultCollection/';
const projectId = '00000000-0000-0000-0000-000000000010';
const repo = {
  id: '00000000-0000-0000-0000-000000000020',
  name: 'RideSharing_Azure_DevOps'
};
const filename = hooks.buildPipelineFilename({
  projectName: 'RideSharing',
  repositoryName: 'RideSharing_Backend',
  branchName: 'feature/defineZones'
});
assert.strictEqual(filename, 'ridesharing-ridesharing_backend-feature-definezones.yml');
assert.strictEqual(hooks.buildPipelineName(filename), filename);
assert.strictEqual(hooks.getAuthHeader('extension-session-token'), 'Bearer extension-session-token');
assert.strictEqual(
  hooks.getDialogConfiguration({ getConfiguration: () => ({ projectId, branch: 'feature/defineZones' }) }).branch,
  'feature/defineZones'
);
assert.strictEqual(
  hooks.getDialogConfiguration({ getConfiguration: () => ({ configuration: { projectId } }) }).projectId,
  projectId
);
assert.strictEqual(
  hooks.buildSignOutUrl('https://azure.example.local/DefaultCollection/'),
  'https://azure.example.local/DefaultCollection/_signout'
);
assert.strictEqual(
  hooks.buildExtensionManagementUrl('https://azure.example.local/DefaultCollection/'),
  'https://azure.example.local/DefaultCollection/_settings/extensions?tab=installed'
);
const hostAuthorizationMessage = hooks.normalizeAccessTokenError({
  message: 'Host authorization was not found (HostAuthorizationNotFound).'
});
assert(hooks.isHostAuthorizationError(hostAuthorizationMessage));
assert(hooks.buildTokenRecoveryMessage(hostAuthorizationMessage).includes('Open extension authorization'));
assert(!hooks.buildTokenRecoveryMessage(hostAuthorizationMessage).includes('Sign out and authenticate again'));
hooks.state.projectName = 'RideSharing';

const desiredBuildDefinition = {
  id: 344,
  name: filename,
  path: '\\KOMODO',
  process: { type: 2, yamlFilename: `/${filename}` },
  repository: {
    id: repo.id,
    name: repo.name,
    type: 'TfsGit',
    defaultBranch: 'refs/heads/main'
  }
};

const reuseCalls = [];
const run = async () => {
  const navigationState = await hooks.getHostNavigationState({
    ServiceIds: { Navigation: 'navigation-service' },
    async getService(serviceId) {
      assert.strictEqual(serviceId, 'navigation-service');
      return {
        getCurrentState() {
          return { branch: 'feature/defineZones', projectId };
        }
      };
    }
  });
  assert.strictEqual(navigationState.branch, 'feature/defineZones');
  assert.strictEqual(navigationState.projectId, projectId);

  let navigatedTo;
  hooks.state.hostUri = hostUri;
  hooks.state.accessToken = 'extension-session-token';
  hooks.state.accessTokenError = 'stale error';
  hooks.state.sdk = {
    ServiceIds: { Navigation: 'navigation-service' },
    async getService(serviceId) {
      assert.strictEqual(serviceId, 'navigation-service');
      return {
        navigate(url) {
          navigatedTo = url;
        }
      };
    }
  };
  await hooks.openExtensionAuthorization();
  assert.strictEqual(navigatedTo, `${hostUri}_settings/extensions?tab=installed`);
  assert.strictEqual(hooks.state.accessToken, null);
  assert.strictEqual(hooks.state.accessTokenError, null);

  context.fetch = async (url) =>
    response({ status: 401, body: { message: 'TF400813: The user is not authorized.' }, url });
  await assert.rejects(
    () =>
      hooks.resolveReleaseAgentQueue({
        hostUri,
        projectId,
        queueName: 'PublishDockerAgent',
        accessToken: 'extension-session-token'
      }),
    (error) => error.status === 401 && error.domain === 'release' && error.requiredExtensionScope === 'vso.agentpools'
  );
  await assert.rejects(
    () =>
      hooks.resolveReleaseVariableGroup({
        hostUri,
        projectId,
        groupName: 'KomodoAPI',
        requiredVariableNames: ['AZP_TOKEN', 'KOMODO_API_KEY', 'KOMODO_API_SECRET'],
        accessToken: 'extension-session-token'
      }),
    (error) =>
      error.status === 401 &&
      error.domain === 'release' &&
      error.requiredExtensionScope === 'vso.variablegroups_read'
  );

  let packagedScriptRequest;
  context.fetch = async (url, options = {}) => {
    packagedScriptRequest = { url, options };
    return response({ body: '#!/usr/bin/env bash\necho packaged-script', url });
  };
  const packagedScript = await hooks.resolveReleaseInlineScript({
    releaseConfig: { scriptSource: { type: 'packagedFile', path: 'release-inline-task.sh' } },
    hostUri,
    accessToken: 'extension-session-token'
  });
  assert.strictEqual(packagedScript, '#!/usr/bin/env bash\necho packaged-script');
  assert.strictEqual(packagedScriptRequest.url, 'https://azure.example.local/extension/dist/release-inline-task.sh');
  assert.strictEqual(packagedScriptRequest.options.headers, undefined);

  hooks.state.accessToken = 'extension-session-token';
  hooks.state.accessTokenError = 'stale error';
  await hooks.restartAzureDevOpsSession();
  assert.strictEqual(navigatedTo, `${hostUri}_signout`);
  assert.strictEqual(hooks.state.accessToken, null);
  assert.strictEqual(hooks.state.accessTokenError, null);

  context.fetch = async (url, options = {}) => {
    reuseCalls.push({ url, options });
    if (reuseCalls.length === 1) {
      // Azure DevOps Server returns only a sparse Pipeline reference here.
      return response({ body: { value: [{ id: 344, name: filename }] }, url });
    }
    if (reuseCalls.length === 2) {
      assert(url.includes('/_apis/build/definitions/344?'));
      return response({ body: desiredBuildDefinition, url });
    }
    throw new Error(`Unexpected reuse request: ${options.method || 'GET'} ${url}`);
  };
  const reused = await hooks.upsertPipelineDefinition({
    hostUri,
    projectId,
    repo,
    pipelineName: filename,
    pipelinePath: `/${filename}`,
    branch: 'main',
    accessToken: 'test-token'
  });
  assert.strictEqual(reused.id, 344);
  assert.strictEqual(reuseCalls.length, 2);
  assert(reuseCalls.every(({ options }) => !options.method || options.method === 'GET'));

  const yaml = '# generated pipeline\ntrigger: none\n';
  const scaffoldCalls = [];
  context.fetch = async (url, options = {}) => {
    scaffoldCalls.push({ url, options });
    if (url.includes('/refs?')) {
      return response({ body: { value: [{ objectId: '1111111111111111111111111111111111111111' }] }, url });
    }
    if (url.includes('/items?')) {
      assert(url.includes('%24format=text'));
      return response({ body: yaml, url });
    }
    throw new Error(`An unchanged YAML must not be pushed: ${options.method || 'GET'} ${url}`);
  };
  const scaffold = await hooks.postScaffold({
    hostUri,
    projectId,
    repoId: repo.id,
    branch: 'main',
    accessToken: 'test-token',
    content: yaml,
    pipelineFilename: filename
  });
  assert.strictEqual(scaffold.skipped, true);
  assert.strictEqual(scaffold.unchanged, true);
  assert.strictEqual(scaffoldCalls.length, 2);
  assert(scaffoldCalls.every(({ options }) => !options.method || options.method === 'GET'));

  const migrationCalls = [];
  context.fetch = async (url, options = {}) => {
    const method = options.method || 'GET';
    migrationCalls.push({ url, method, options });
    if (url.includes('/_apis/pipelines?')) {
      return response({ body: { value: [] }, url });
    }
    if (url.includes('/_apis/build/definitions?')) {
      return response({
        body: {
          value: [
            {
              id: 345,
              revision: 2,
              name: 'RideSharing_RideSharing_Azure_DevOps_demo',
              path: '\\KOMODO',
              process: { type: 2, yamlFilename: `/${filename}` },
              repository: { id: repo.id, defaultBranch: 'refs/heads/main' }
            },
            {
              id: 344,
              revision: 7,
              name: 'RideSharing_RideSharing_Backend_demo',
              path: '\\KOMODO',
              process: { type: 2, yamlFilename: `/${filename}` },
              repository: { id: repo.id, defaultBranch: 'refs/heads/main' }
            }
          ]
        },
        url
      });
    }
    if (url.includes('/_apis/build/definitions/344') && method === 'GET') {
      return response({
        body: {
          id: 344,
          revision: 7,
          name: 'RideSharing_RideSharing_Backend_demo',
          path: '\\KOMODO',
          process: { type: 2, yamlFilename: `/${filename}` },
          repository: { id: repo.id, name: repo.name, type: 'TfsGit', defaultBranch: 'refs/heads/main' }
        },
        url
      });
    }
    if (url.includes('/_apis/build/definitions/344') && method === 'PUT') {
      const body = JSON.parse(options.body);
      assert.strictEqual(body.id, 344);
      assert.strictEqual(body.revision, 7);
      assert.strictEqual(body.name, filename);
      assert.strictEqual(body.path, '\\komodo');
      assert.strictEqual(body.process.yamlFilename, `/${filename}`);
      assert.strictEqual(body.repository.id, repo.id);
      return response({ body: { ...body, id: 344, revision: 8 }, url });
    }
    throw new Error(`Unexpected migration request: ${method} ${url}`);
  };

  const migrated = await hooks.upsertPipelineDefinition({
    hostUri,
    projectId,
    repo,
    pipelineName: filename,
    pipelinePath: `/${filename}`,
    branch: 'main',
    accessToken: 'test-token'
  });
  assert.strictEqual(migrated.id, 344);
  assert(migrationCalls.some(({ url, method }) => url.includes('/_apis/build/definitions/344') && method === 'PUT'));
  assert(!migrationCalls.some(({ url, method }) => url.includes('/_apis/pipelines/344') && method === 'PUT'));

  const releaseCalls = [];
  let updatedReleaseBody;
  context.fetch = async (url, options = {}) => {
    const method = options.method || 'GET';
    const parsed = new URL(url);
    releaseCalls.push({ url, method, options });
    if (url.includes('/_apis/release/definitions?') && method === 'GET' && parsed.searchParams.has('searchText')) {
      return response({ body: { value: [] }, url });
    }
    if (url.includes('/_apis/release/definitions?') && method === 'GET' && parsed.searchParams.has('artifactSourceId')) {
      assert.strictEqual(parsed.searchParams.get('artifactSourceId'), `${projectId}:344`);
      assert.strictEqual(parsed.searchParams.get('$expand'), 'Artifacts');
      return response({
        body: {
          value: [
            {
              id: 5,
              name: 'RideSharing_RideSharing_Backend_demo_Release',
              artifacts: [{ definitionReference: { definition: { id: '344' } } }]
            }
          ]
        },
        url
      });
    }
    if (url.includes('/_apis/distributedtask/queues?') && method === 'GET') {
      return response({ body: { value: [{ id: 111, name: 'PublishDockerAgent' }] }, url });
    }
    if (url.includes('/_apis/distributedtask/variablegroups?') && method === 'GET') {
      assert.strictEqual(parsed.searchParams.get('groupName'), 'KomodoAPI');
      assert.strictEqual(parsed.searchParams.get('actionFilter'), 'Use');
      return response({
        body: {
          value: [
            {
              id: 7,
              name: 'KomodoAPI',
              variables: {
                AZP_TOKEN: { isSecret: true, value: null },
                KOMODO_API_KEY: { isSecret: true, value: null },
                KOMODO_API_SECRET: { isSecret: true, value: null }
              }
            }
          ]
        },
        url
      });
    }
    if (url.includes('/_apis/release/definitions/5?') && method === 'GET') {
      return response({
        body: {
          id: 5,
          revision: 4,
          name: 'RideSharing_RideSharing_Backend_demo_Release',
          path: '\\komodo',
          variableGroups: [9],
          artifacts: [{ definitionReference: { definition: { id: '344' }, repository: { id: repo.id } } }],
          environments: [
            {
              id: 23,
              name: 'komodo',
              conditions: [],
              preDeployApprovals: { approvals: [] },
              postDeployApprovals: { approvals: [] },
              deployPhases: [
                {
                  id: 31,
                  workflowTasks: [],
                  deploymentInput: { queueId: 111 }
                }
              ]
            }
          ]
        },
        url
      });
    }
    if (url.includes('/_apis/release/definitions?') && method === 'PUT') {
      const body = JSON.parse(options.body);
      assert.strictEqual(body.id, 5);
      assert.strictEqual(body.revision, 4);
      assert.strictEqual(body.name, `${filename}_Release`);
      assert.strictEqual(body.environments[0].id, 23);
      assert.strictEqual(body.environments[0].deployPhases[0].id, 31);
      assert.strictEqual(body.artifacts[0].definitionReference.definition.id, '344');
      assert.deepStrictEqual(body.variableGroups, [9, 7]);
      assert.strictEqual(
        body.environments[0].deployPhases[0].workflowTasks[0].inputs.script,
        '#!/usr/bin/env bash\necho regression-test'
      );
      updatedReleaseBody = body;
      return response({ body: { ...body, id: 5, revision: 5 }, url });
    }
    throw new Error(`Unexpected Release request: ${method} ${url}`);
  };

  const release = await hooks.ensureReleaseDefinition({
    hostUri,
    projectId,
    projectName: 'RideSharing',
    repo,
    pipelineDefinition: { id: 344 },
    pipelineName: filename,
    branch: 'main',
    queueName: 'PublishDockerAgent',
    accessToken: 'test-token'
  });
  assert.strictEqual(release.id, 5);
  assert.strictEqual(release.created, false);
  assert.strictEqual(release.updated, true);
  assert(releaseCalls.some(({ method, url }) => method === 'PUT' && url.includes('/_apis/release/definitions?')));

  const noOpReleaseCalls = [];
  context.fetch = async (url, options = {}) => {
    const method = options.method || 'GET';
    const parsed = new URL(url);
    noOpReleaseCalls.push({ url, method });
    if (url.includes('/_apis/release/definitions?') && method === 'GET' && parsed.searchParams.has('searchText')) {
      return response({ body: { value: [{ id: 5, name: `${filename}_Release` }] }, url });
    }
    if (url.includes('/_apis/distributedtask/queues?') && method === 'GET') {
      return response({ body: { value: [{ id: 111, name: 'PublishDockerAgent' }] }, url });
    }
    if (url.includes('/_apis/distributedtask/variablegroups?') && method === 'GET') {
      return response({
        body: {
          value: [
            {
              id: 7,
              name: 'KomodoAPI',
              variables: {
                AZP_TOKEN: { isSecret: true, value: null },
                KOMODO_API_KEY: { isSecret: true, value: null },
                KOMODO_API_SECRET: { isSecret: true, value: null }
              }
            }
          ]
        },
        url
      });
    }
    if (url.includes('/_apis/release/definitions/5?') && method === 'GET') {
      return response({ body: { ...updatedReleaseBody, id: 5, revision: 5 }, url });
    }
    throw new Error(`A matching Release must not be written: ${method} ${url}`);
  };
  const reusedRelease = await hooks.ensureReleaseDefinition({
    hostUri,
    projectId,
    projectName: 'RideSharing',
    repo,
    pipelineDefinition: { id: 344 },
    pipelineName: filename,
    branch: 'main',
    queueName: 'PublishDockerAgent',
    accessToken: 'test-token'
  });
  assert.strictEqual(reusedRelease.id, 5);
  assert.strictEqual(reusedRelease.created, false);
  assert.strictEqual(reusedRelease.updated, false);
  assert(noOpReleaseCalls.every(({ method }) => method === 'GET'));

  console.log(
    'UI behavior regression tests passed: session restart URL, exact naming, unchanged YAML and canonical Build/Release no-op reuse, no Pipelines PUT, and Pipeline/Release/KomodoAPI reconciliation.'
  );
};

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
