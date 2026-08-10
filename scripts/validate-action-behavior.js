#!/usr/bin/env node

/*
 * Focused regression test for the branch action launch path. The production
 * source is instrumented in memory; no test hook is shipped and no network or
 * Azure DevOps call is made.
 */
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const actionPath = path.join(root, 'dist/menu-action.js');
const source = fs.readFileSync(actionPath, 'utf8');
const initializationMarker = '\ninitializeAction();\n';
if (!source.includes(initializationMarker)) {
  throw new Error('[action behavior validation] Could not find the action initialization marker.');
}

const instrumented = source.replace(
  initializationMarker,
  '\nwindow.__PipelineGeneratorActionTestHooks = { openGenerator, buildDialogContributionId, buildGeneratorHubUrl };\n'
);

const dialogCalls = [];
const navigationCalls = [];
let tokenCalls = 0;
let detachedWindowCalls = 0;
let dialogUnavailable = false;
const hostUri = 'https://azure.example.local/DefaultCollection/';
const project = { id: '00000000-0000-0000-0000-000000000010', name: 'RideSharing' };
const repository = {
  id: '00000000-0000-0000-0000-000000000020',
  name: 'RideSharing_Backend',
  project
};

const hostPageLayoutService = {
  openCustomDialog(contributionId, options) {
    if (dialogUnavailable) {
      throw new Error('Host dialog service is unavailable on this Azure DevOps Server.');
    }
    dialogCalls.push({ contributionId, options });
  }
};

const hostNavigationService = {
  navigate(url) {
    navigationCalls.push(url);
  }
};

const sdk = {
  ServiceIds: { Navigation: 'ms.vss-web.navigation-service' },
  async getService(serviceId) {
    if (serviceId === 'ms.vss-features.host-page-layout-service') return hostPageLayoutService;
    if (serviceId === 'ms.vss-web.navigation-service') return hostNavigationService;
    throw new Error(`Unexpected service ID: ${serviceId}`);
  },
  async getAccessToken() {
    tokenCalls += 1;
    return 'must-not-be-requested-on-dialog-path';
  }
};

const VSS = {
  getConfiguration: () => ({}),
  getWebContext: () => ({ collection: { uri: hostUri }, project, repository }),
  getExtensionContext: () => ({
    publisherId: 'mohammad-falahat',
    extensionId: 'pipeline-generator',
    baseUri: 'https://azure.example.local/_apis/public/gallery/publisher/mohammad-falahat/extension/pipeline-generator/0.1.26/assetbyname/'
  })
};

const window = {
  VSS,
  location: {
    origin: 'https://azure.example.local',
    href: `${hostUri}RideSharing/_git/RideSharing_Backend/branches?_a=all`
  },
  open() {
    detachedWindowCalls += 1;
  },
  addEventListener() {},
  removeEventListener() {}
};
window.parent = window;

const document = {
  referrer: `${hostUri}RideSharing/_git/RideSharing_Backend/branches?_a=all`,
  createElement: () => ({}) ,
  head: { appendChild() {} }
};

const quietConsole = { log() {}, warn() {}, error() {} };
const context = {
  VSS,
  window,
  document,
  URL,
  URLSearchParams,
  setTimeout,
  clearTimeout,
  setInterval,
  clearInterval,
  console: quietConsole,
  fetch: async () => {
    throw new Error('Unexpected fetch call.');
  }
};

vm.runInNewContext(instrumented, context, { filename: 'dist/menu-action.js' });
const hooks = window.__PipelineGeneratorActionTestHooks;
assert(hooks, 'Action test hooks were not exposed by the in-memory instrumentation.');
assert.strictEqual(
  hooks.buildDialogContributionId(VSS.getExtensionContext()),
  'mohammad-falahat.pipeline-generator.pipeline-generator-dialog'
);

const run = async () => {
  await hooks.openGenerator(
    {
      gitRepository: repository,
      project,
      branch: { name: 'refs/heads/feature/defineZones', repository }
    },
    sdk
  );

  assert.strictEqual(dialogCalls.length, 1);
  assert.strictEqual(detachedWindowCalls, 0);
  assert.strictEqual(tokenCalls, 0, 'The action must not request a token before opening the host dialog.');

  const [{ contributionId, options }] = dialogCalls;
  assert.strictEqual(contributionId, 'mohammad-falahat.pipeline-generator.pipeline-generator-dialog');
  assert.strictEqual(options.lightDismiss, false);
  assert.strictEqual(options.configuration.branch, 'feature/defineZones');
  assert.strictEqual(options.configuration.projectId, project.id);
  assert.strictEqual(options.configuration.repoId, repository.id);
  assert.strictEqual(options.configuration.hostUri, hostUri);
  assert(!Object.prototype.hasOwnProperty.call(options.configuration, 'accessToken'));
  assert(!Object.prototype.hasOwnProperty.call(options.configuration, 'accessTokenError'));

  dialogUnavailable = true;
  await hooks.openGenerator(
    {
      gitRepository: repository,
      project,
      branch: { name: 'refs/heads/feature/defineZones', repository }
    },
    sdk
  );

  assert.strictEqual(navigationCalls.length, 1);
  assert.strictEqual(detachedWindowCalls, 0);
  assert.strictEqual(tokenCalls, 0, 'The action must not request a token for the in-host hub fallback.');
  const hubUrl = new URL(navigationCalls[0]);
  assert.strictEqual(
    hubUrl.pathname,
    '/DefaultCollection/RideSharing/_apps/hub/mohammad-falahat.pipeline-generator.pipeline-generator-hub'
  );
  assert.strictEqual(hubUrl.searchParams.get('branch'), 'feature/defineZones');
  assert.strictEqual(hubUrl.searchParams.get('projectId'), project.id);
  assert.strictEqual(hubUrl.searchParams.get('repoId'), repository.id);
  assert.strictEqual(hubUrl.searchParams.get('hostUri'), hostUri);

  console.log('Action behavior regression test passed: dialog or in-host Repos hub opens with token-free context.');
};

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
