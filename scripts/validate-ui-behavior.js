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
const yaml = require('js-yaml');

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
	    buildLegacyServiceLessPipelineFilename,
	    buildLegacyPipelineFilename,
	    buildLegacyEnvironmentFirstPipelineFilename,
	    buildPipelineName,
	    buildReleaseName,
	    normalizeGeneratorMode,
	    applyModePresentation,
	    getPipelineFolder,
	    getReleaseConfig,
	    buildMonorepoDeploymentContract,
	    buildMonorepoPipelineYaml,
	    buildMonorepoSupportRepositorySpecs,
	    buildMonorepoComposeSample,
	    mergeMonorepoComposeServices,
	    buildMonorepoNginxSample,
	    mergeMonorepoNginxRoutes,
	    buildCollectionUri,
	    buildCentralGitItemUrl,
	    parseDeploymentTargetsYaml,
	    fetchDeploymentTargets,
	    parseKomodoCredentialFile,
	    fetchKomodoCredentials,
	    extractEnabledKomodoServers,
	    fetchKomodoServers,
	    extractDockerNetworkNames,
	    selectNginxNetworkName,
	    fetchKomodoDockerNetworks,
	    resolveNginxNetworkForServer,
	    loadDeploymentTargets,
	    setKomodoServerFromEnvironment,
	    setServiceNameFromRepository,
	    buildSupportRepositorySpecs,
	    buildComposeSample,
	    buildNginxRouteBlock,
	    buildNginxSample,
	    mergeNginxServiceRoute,
	    ensureSupportRepositories,
	    ensureRepositoryBootstrapFiles,
	    buildRepositoryFileUrl,
	    showCompletionLinks,
	    finishProvisioning,
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
	    postGeneratedFiles,
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
  value: '',
  options: []
});
const elements = new Map([
  ['branch-label', element()],
  ['page-title', element()],
  ['form-hint', element()],
  ['branch', element()],
  ['environment', environment],
  ['pool', element()],
  ['service', element()],
  ['containerRegistryService', element()],
  ['repositoryAddress', element()],
  ['dockerfileDir', element()],
  ['pipeline-form', form],
  ['status', element()],
  ['targetRepo', element()],
  ['komodoServer', element({ options: [] })],
  ['reauth-panel', element({ className: 'auth-fallback hidden' })],
  ['reauth-message', element()],
  ['authorize-extension', element()],
  ['reauthenticate', element()],
  ['completion-panel', element({ className: 'completion-panel hidden' })],
  ['nginx-result-link', element()],
  ['compose-result-link', element()],
  ['pipeline-result-link', element()],
  ['contract-result-item', element({ className: 'hidden' })],
  ['contract-result-link', element()],
  ['completion-hint', element()],
  ['service-field', element()],
  ['dockerfile-field', element()],
  ['registry-address-field', element()],
  ['registry-service-field', element()]
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
  service: 'api',
  environment: 'demo',
  branchName: 'feature/defineZones'
});
const serviceLessFilename = hooks.buildLegacyServiceLessPipelineFilename({
  projectName: 'RideSharing',
  repositoryName: 'RideSharing_Backend',
  environment: 'demo',
  branchName: 'feature/defineZones'
});
const previousEnvironmentFirstFilename = hooks.buildLegacyEnvironmentFirstPipelineFilename({
  projectName: 'RideSharing',
  repositoryName: 'RideSharing_Backend',
  environment: 'demo',
  branchName: 'feature/defineZones'
});
const legacyFilename = hooks.buildLegacyPipelineFilename({
  projectName: 'RideSharing',
  repositoryName: 'RideSharing_Backend',
  branchName: 'feature/defineZones'
});
assert.strictEqual(filename, 'ridesharing-ridesharing_backend-api-Feature-DefineZonesToDEMO.yml');
assert.strictEqual(serviceLessFilename, 'ridesharing-ridesharing_backend-Feature-DefineZonesToDEMO.yml');
assert.strictEqual(previousEnvironmentFirstFilename, 'ridesharing-ridesharing_backend-demo-feature-definezones.yml');
assert.strictEqual(legacyFilename, 'ridesharing-ridesharing_backend-feature-definezones.yml');
assert.strictEqual(hooks.buildPipelineName(filename), filename);
assert.strictEqual(hooks.buildReleaseName({ service: 'api', environment: 'demo' }), 'API DEMO');
assert.notStrictEqual(
  hooks.buildReleaseName({ service: 'api', environment: 'demo' }),
  hooks.buildReleaseName({ service: 'worker', environment: 'demo' })
);
assert.strictEqual(
  hooks.buildPipelineFilename({
    projectName: 'RideSharing',
    repositoryName: 'RideSharing_Backend',
    service: 'frontend',
    environment: 'demo',
    branchName: 'feature/defineZones',
    mode: 'monorepo'
  }),
  'ridesharing-ridesharing_backend-MR-frontend-Feature-DefineZonesToDEMO.yml'
);
assert.strictEqual(
  hooks.buildReleaseName({ service: 'frontend', environment: 'demo', mode: 'monorepo' }),
  'MR FRONTEND DEMO'
);
assert.notStrictEqual(
  filename,
  hooks.buildPipelineFilename({
    projectName: 'RideSharing',
    repositoryName: 'RideSharing_Backend',
    service: 'worker',
    environment: 'demo',
    branchName: 'feature/defineZones'
  })
);
hooks.applyModePresentation('monorepo');
assert.strictEqual(hooks.state.mode, 'monorepo');
assert.strictEqual(hooks.getPipelineFolder(), '\\komodo\\MR');
assert.strictEqual(elements.get('service-field').hidden, false);
assert.strictEqual(elements.get('dockerfile-field').hidden, true);
assert.strictEqual(submitButton.textContent, 'Create MR runtime, pipeline, and release');
const monorepoReleaseConfig = hooks.getReleaseConfig('monorepo');
assert.strictEqual(monorepoReleaseConfig.folder, '\\komodo\\MR');
assert.strictEqual(monorepoReleaseConfig.scriptSource.path, 'monorepo-release-inline-task.sh');
assert.strictEqual(monorepoReleaseConfig.bashTaskName, 'Deploy immutable MR images through Komodo');
assert.strictEqual(elements.get('registry-address-field').hidden, false);
assert.strictEqual(elements.get('registry-service-field').hidden, false);
assert.strictEqual(elements.get('repositoryAddress').required, true);
assert.strictEqual(elements.get('containerRegistryService').required, true);
hooks.applyModePresentation('pipeline');
assert.strictEqual(hooks.getPipelineFolder(), '\\komodo');
assert.strictEqual(elements.get('service-field').hidden, false);

const monorepoContract = hooks.buildMonorepoDeploymentContract();
assert(monorepoContract.includes('kind: nx-monorepo'));
assert(monorepoContract.includes('install_command: "pnpm install --frozen-lockfile"'));
assert(monorepoContract.includes('continue_on_module_error: true'));
assert(monorepoContract.includes('orphan_policy: "retain"'));
const monorepoYaml = hooks.buildMonorepoPipelineYaml(
  {
    pool: 'PublishDockerAgent',
    environment: 'demo',
    komodoServer: 'DEMO-192.168.62.91',
    repositoryAddress: 'registry.buluttakin.com',
    containerRegistryService: 'BulutReg'
  },
  {
    sourceBranch: 'feature/defineZones',
    rawProjectName: 'RideSharing',
    rawRepositoryName: 'RideSharing_FrontEnd'
  }
);
assert.doesNotThrow(() => yaml.load(monorepoYaml));
assert(monorepoYaml.includes('name: \'RideSharing/RideSharing_FrontEnd\''));
assert(monorepoYaml.includes('repository: SharedTemplatesRepo'));
assert(monorepoYaml.includes('endpoint: ShonizCollection'));
assert(monorepoYaml.includes('template: monorepo/pipeline.yml@SharedTemplatesRepo'));
assert(monorepoYaml.includes('- group: KomodoAPI'));
assert(monorepoYaml.includes("environment: 'demo'"));
assert(monorepoYaml.includes("serviceKey: 'frontend'"));
assert(monorepoYaml.includes("komodoServer: 'DEMO-192.168.62.91'"));
assert(!monorepoYaml.includes('deploymentRoot:'));
assert(monorepoYaml.includes("composeRepository: 'RideSharing_Docker_DevOps'"));
assert(monorepoYaml.includes("composePath: '/demo_ridesharing/compose.yml'"));
assert(monorepoYaml.includes("komodoRepository: 'RideSharing_Docker_DevOps-demo'"));
assert(monorepoYaml.includes("komodoStack: 'RideSharing_Docker_DevOps-demo'"));
assert(monorepoYaml.includes("staticContainer: 'ridesharing_frontend_demo'"));
assert(monorepoYaml.includes("bffContainer: 'ridesharing_frontend_bff_demo'"));
assert(monorepoYaml.includes("bffProfile: 'mr-frontend-bff'"));
assert(monorepoYaml.includes("registryAddress: 'registry.buluttakin.com'"));
assert(monorepoYaml.includes("containerRegistryService: 'BulutReg'"));
assert(monorepoYaml.includes("staticRuntimeImage: 'registry.buluttakin.com/nginx:1.27-alpine'"));
assert(monorepoYaml.includes("bffRuntimeImage: 'registry.buluttakin.com/node:20-alpine'"));
assert(monorepoYaml.includes("nodeImage: 'registry.buluttakin.com/node:22-bookworm'"));
assert(!monorepoYaml.includes('/.devops/mr-build.cjs'));
assert(!monorepoYaml.includes('PublishBuildArtifacts@1'));
assert.strictEqual(
  hooks.buildPipelineFilename({
    projectName: 'RideSharing',
    repositoryName: 'RideSharing_Backend',
    service: 'api',
    environment: 'dev',
    branchName: 'feature/defineZones'
  }),
  'ridesharing-ridesharing_backend-api-Feature-DefineZonesToDEV.yml'
);
assert.strictEqual(
  hooks.buildPipelineFilename({
    projectName: 'Locanit',
    repositoryName: 'Locanit_API',
    service: 'api',
    environment: 'soc',
    branchName: 'Production'
  }),
  'locanit-locanit_api-api-ProductionToSOC.yml'
);
assert.throws(
  () =>
    hooks.buildPipelineFilename({
      projectName: 'RideSharing',
      repositoryName: 'RideSharing_Backend',
      service: 'api',
      branchName: 'feature/defineZones'
    }),
  /Environment is required/
);
assert.throws(
  () =>
    hooks.buildPipelineFilename({
      projectName: 'RideSharing',
      repositoryName: 'RideSharing_Backend',
      environment: 'demo',
      branchName: 'feature/defineZones'
    }),
  /Service name is required/
);
hooks.setServiceNameFromRepository('Locanit_API', 'Locanit');
assert.strictEqual(elements.get('service').value, 'api');
const deploymentTargets = hooks.parseDeploymentTargetsYaml(`
servers:
  - "DEMO-192.168.62.91"
  - "Development-192.168.62.19"
  - "Production-192.168.0.244"
  - "Production-192.168.62.140"
  - "Production-31.7.65.195"
  - "QA-192.168.62.153"
environments:
  - name: pro
    domain: bulutcom.cloud
  - name: qa
    domain: bulutqa.ir
  - name: demo # inline comments are allowed
    domain: bulutdemo.ir
  - "dev:bulutdev.ir" # compact legacy form remains accepted
  - name: soc
    domain: bulutsoc.ir
`);
assert.deepStrictEqual(Array.from(deploymentTargets.servers), [
  'DEMO-192.168.62.91',
  'Development-192.168.62.19',
  'Production-192.168.0.244',
  'Production-192.168.62.140',
  'Production-31.7.65.195',
  'QA-192.168.62.153'
]);
assert.deepStrictEqual(Array.from(deploymentTargets.environments), ['pro', 'qa', 'demo', 'dev', 'soc']);
assert.deepStrictEqual(
  Array.from(deploymentTargets.environmentConfigs, (item) => ({ ...item })),
  [
    { name: 'pro', domain: 'bulutcom.cloud', projectsRoot: '/mnt/graid/projects' },
    { name: 'qa', domain: 'bulutqa.ir', projectsRoot: '/var/data/projects' },
    { name: 'demo', domain: 'bulutdemo.ir', projectsRoot: '/var/data/projects' },
    { name: 'dev', domain: 'bulutdev.ir', projectsRoot: '/var/data/projects' },
    { name: 'soc', domain: 'bulutsoc.ir', projectsRoot: '/var/data/projects' }
  ]
);
assert.throws(
  () => hooks.parseDeploymentTargetsYaml('environments:\n  - dev\n'),
  /must define a valid domain/
);
assert.deepStrictEqual(
  Array.from(hooks.extractDockerNetworkNames([
    { name: 'bridge' },
    { name: 'nginx-net' },
    { name: 'nginx-net' }
  ])),
  ['bridge', 'nginx-net']
);
assert.strictEqual(hooks.selectNginxNetworkName(['bridge', 'nginx-net']), 'nginx-net');
assert.strictEqual(hooks.selectNginxNetworkName(['nginx-network', 'nginx-net']), 'nginx-network');
assert.throws(
  () => hooks.selectNginxNetworkName(['bridge', 'nginx-net-demo']),
  /neither nginx-network nor nginx-net/
);
const komodoServerSelect = elements.get('komodoServer');
komodoServerSelect.options = deploymentTargets.servers.map((value) => ({ value }));
hooks.setKomodoServerFromEnvironment('dev');
assert.strictEqual(komodoServerSelect.value, 'Development-192.168.62.19');
hooks.setKomodoServerFromEnvironment('soc');
assert.strictEqual(komodoServerSelect.value, '');
assert.deepStrictEqual(
  Array.from(
    hooks.buildSupportRepositorySpecs({
      projectName: '180 Feedback',
      environment: 'demo',
      domain: 'bulutdemo.ir',
      service: 'api',
      repositoryAddress: 'registry.buluttakin.com'
    }),
    (item) => ({ name: item.name, directory: item.directory, filePath: item.filePath })
  ),
  [
    {
      name: '180Feedback_Docker_DevOps',
      directory: 'demo_180feedback',
      filePath: '/demo_180feedback/compose.yml'
    },
    {
      name: '180Feedback_Nginx_DevOps',
      directory: 'demo',
      filePath: '/demo/180feedback-demo.conf'
    }
  ]
);
const monorepoSpecs = hooks.buildMonorepoSupportRepositorySpecs({
  projectName: '180 Feedback',
  environment: 'demo',
  domain: 'bulutdemo.ir',
  service: 'frontend'
});
assert.deepStrictEqual(
  Array.from(monorepoSpecs, (item) => ({ name: item.name, directory: item.directory, filePath: item.filePath })),
  [
    {
      name: '180Feedback_Docker_DevOps',
      directory: 'demo_180feedback',
      filePath: '/demo_180feedback/compose.yml'
    },
    {
      name: '180Feedback_Nginx_DevOps',
      directory: 'demo',
      filePath: '/demo/180feedback-demo.conf'
    }
  ]
);
const monorepoDockerSpec = monorepoSpecs.find((item) => item.kind === 'docker');
const monorepoCompose = monorepoDockerSpec.content;
const monorepoEnvSpec = monorepoDockerSpec.additionalFiles.find((item) => item.path.endsWith('/.env'));
assert.doesNotThrow(() => yaml.load(monorepoCompose));
assert(!monorepoCompose.includes('name: 180feedback-mr-demo'));
assert(monorepoCompose.includes('container_name: 180feedback_frontend_demo'));
assert(monorepoCompose.includes('container_name: 180feedback_frontend_bff_demo'));
assert(monorepoCompose.includes('image: registry.buluttakin.com/180feedback/frontend-demo:${frontend}'));
assert(monorepoCompose.includes('image: registry.buluttakin.com/180feedback/frontend-bff-demo:${frontend_bff}'));
assert.strictEqual(monorepoEnvSpec.path, '/demo_180feedback/.env');
assert(monorepoEnvSpec.content.includes('frontend:CHANGE_ME'));
assert(monorepoEnvSpec.content.includes('frontend_bff:CHANGE_ME'));
assert.strictEqual(
  monorepoEnvSpec.mergeExisting('api:1.0.100\nfrontend:1.0.14999\n'),
  'api:1.0.100\nfrontend:1.0.14999\nfrontend_bff:CHANGE_ME\n'
);
assert(!monorepoCompose.includes('volumes:'));
assert(!monorepoCompose.includes('cat > /etc/nginx/conf.d/default.conf'));
assert(monorepoCompose.includes('profiles: ["mr-frontend-bff"]'));
assert(!monorepoCompose.includes('working_dir:'));
assert(!monorepoCompose.includes('command:'));
assert(!monorepoCompose.includes('/var/data/projects'));
assert(!monorepoCompose.includes('/mnt/graid/projects'));
const productionMonorepoCompose = hooks.buildMonorepoComposeSample({
  projectKey: '180feedback',
  serviceKey: 'frontend',
  environment: 'pro'
});
assert(!productionMonorepoCompose.includes('/mnt/graid/projects'));
assert(!monorepoCompose.includes('image: registry.buluttakin.com/nginx:1.27-alpine'));
assert(!monorepoCompose.includes('image: registry.buluttakin.com/node:20-alpine'));
const existingCompose = [
  'services:',
  '  180feedback_api_demo:',
  '    image: registry.example/api:123',
  '    networks:',
  '      - nginx-network',
  '',
  'networks:',
  '  nginx-network:',
  '    external: true',
  ''
].join('\n');
const mergedMonorepoCompose = hooks.mergeMonorepoComposeServices({
  content: existingCompose,
  compactProject: '180Feedback',
  projectKey: '180feedback',
  serviceKey: 'frontend',
  environment: 'demo'
});
assert.doesNotThrow(() => yaml.load(mergedMonorepoCompose));
assert(mergedMonorepoCompose.includes('180feedback_api_demo:'));
assert(mergedMonorepoCompose.includes('180feedback_frontend_demo:'));
assert(mergedMonorepoCompose.includes('180feedback_frontend_bff_demo:'));
assert.strictEqual(
  hooks.mergeMonorepoComposeServices({
    content: mergedMonorepoCompose,
    compactProject: '180Feedback',
    projectKey: '180feedback',
    serviceKey: 'frontend',
    environment: 'demo'
  }),
  mergedMonorepoCompose
);
const immutableTaggedCompose = mergedMonorepoCompose.replace(
  'image: registry.buluttakin.com/180feedback/frontend-demo:${frontend}',
  'image: registry.buluttakin.com/180feedback/frontend-demo:1.0.14999'
);
assert(
  hooks.mergeMonorepoComposeServices({
    content: immutableTaggedCompose,
    projectKey: '180feedback',
    serviceKey: 'frontend',
    environment: 'demo',
    repositoryAddress: 'registry.buluttakin.com'
  }).includes('image: registry.buluttakin.com/180feedback/frontend-demo:1.0.14999')
);
const alternateRegistryCompose = hooks.buildMonorepoComposeSample({
  projectKey: '180feedback',
  serviceKey: 'frontend',
  environment: 'demo',
  repositoryAddress: 'registry.internal.example'
});
assert(alternateRegistryCompose.includes('image: registry.internal.example/180feedback/frontend-demo:${frontend}'));
assert(alternateRegistryCompose.includes('image: registry.internal.example/180feedback/frontend-bff-demo:${frontend_bff}'));
const alternateNginxNetworkCompose = hooks.mergeMonorepoComposeServices({
  content: existingCompose,
  projectKey: '180feedback',
  serviceKey: 'frontend',
  environment: 'demo',
  nginxNetworkName: 'nginx-net'
});
assert(alternateNginxNetworkCompose.includes('      - nginx-network'));
assert(alternateNginxNetworkCompose.includes('  nginx-network:\n    name: nginx-net\n    external: true'));
assert.strictEqual(
  hooks.mergeMonorepoComposeServices({
    content: alternateNginxNetworkCompose,
    projectKey: '180feedback',
    serviceKey: 'frontend',
    environment: 'demo',
    nginxNetworkName: 'nginx-net'
  }),
  alternateNginxNetworkCompose
);
const newNginxNetCompose = hooks.buildMonorepoComposeSample({
  projectKey: '180feedback',
  serviceKey: 'frontend',
  environment: 'demo',
  nginxNetworkName: 'nginx-net'
});
assert(newNginxNetCompose.includes('      - nginx-net'));
assert(newNginxNetCompose.includes('  nginx-net:\n    name: nginx-net\n    external: true'));
const legacyManagedImageCompose = mergedMonorepoCompose
  .replace('image: registry.example/api:123', 'image: nginx:1.27-alpine')
  .replace('image: registry.buluttakin.com/180feedback/frontend-demo:${frontend}', 'image: nginx:1.27-alpine')
  .replace('image: registry.buluttakin.com/180feedback/frontend-bff-demo:${frontend_bff}', 'image: node:20-alpine')
  .replace(
    '    restart: unless-stopped\n    expose:',
    '    restart: unless-stopped\n    volumes:\n      - /mnt/graid/projects/180Feedback_Docker_DevOps/demo_180feedback/monorepo/frontend:/srv/monorepo:ro\n      - /mnt/graid/projects/180Feedback_Docker_DevOps/demo_180feedback/monorepo/frontend/runtime/nginx/default.conf:/etc/nginx/conf.d/default.conf:ro\n    expose:'
  )
  .replace(
    '    profiles: ["mr-frontend-bff"]\n    restart: unless-stopped',
    '    profiles: ["mr-frontend-bff"]\n    restart: unless-stopped\n    working_dir: /srv/monorepo/current/modules/${MR_180FEEDBACK_FRONTEND_DEMO_BFF_PROJECT:-bff}\n    command: ["/bin/sh", "-ec", "exec node \\"${MR_180FEEDBACK_FRONTEND_DEMO_BFF_ENTRY:-main.js}\\""]\n    volumes:\n      - /var/data/projects/180Feedback_Docker_DevOps/demo_180feedback/monorepo/frontend:/srv/monorepo:ro'
  );
const migratedManagedImageCompose = hooks.mergeMonorepoComposeServices({
  content: legacyManagedImageCompose,
  compactProject: '180Feedback',
  projectKey: '180feedback',
  serviceKey: 'frontend',
  environment: 'demo'
});
assert(migratedManagedImageCompose.includes('180feedback_api_demo:\n    image: nginx:1.27-alpine'));
assert(migratedManagedImageCompose.includes('image: registry.buluttakin.com/180feedback/frontend-demo:${frontend}'));
assert(migratedManagedImageCompose.includes('image: registry.buluttakin.com/180feedback/frontend-bff-demo:${frontend_bff}'));
assert(!migratedManagedImageCompose.includes('/var/data/projects'));
assert(!migratedManagedImageCompose.includes('/mnt/graid/projects'));
assert(!migratedManagedImageCompose.includes('working_dir:'));
assert(!migratedManagedImageCompose.includes('command:'));
assert(!migratedManagedImageCompose.includes('volumes:'));
assert.strictEqual(
  hooks.mergeMonorepoComposeServices({
    content: migratedManagedImageCompose,
    compactProject: '180Feedback',
    projectKey: '180feedback',
    serviceKey: 'frontend',
    environment: 'demo'
  }),
  migratedManagedImageCompose
);
const monorepoNginx = monorepoSpecs.find((item) => item.kind === 'nginx').content;
assert(monorepoNginx.includes('server_name 180feedback.bulutdemo.ir;'));
assert(monorepoNginx.includes('/etc/nginx/conf.d/bulutdemo.ir.pem'));
assert(monorepoNginx.includes('/etc/nginx/conf.d/bulutdemo.ir.key'));
assert(monorepoNginx.includes('location /bff/ {'));
assert(monorepoNginx.includes('set              $target            180feedback_frontend_bff_demo;'));
assert(monorepoNginx.includes('proxy_pass                          http://$target:3000;'));
assert(monorepoNginx.includes('location / {'));
assert(monorepoNginx.includes('set              $target            180feedback_frontend_demo;'));
assert(monorepoNginx.includes('proxy_pass                          http://$target:80;'));
assert(!monorepoNginx.includes('proxy_pass                          http://$target:80/;'));
assert(!monorepoNginx.includes('rewrite '));
assert(monorepoNginx.indexOf('location /bff/ {') < monorepoNginx.indexOf('location / {'));
const legacyMonorepoNginx = monorepoNginx
  .replace('/etc/nginx/conf.d/bulutdemo.ir.pem', '"/etc/nginx/conf.d/bulutdemo.pem"')
  .replace('/etc/nginx/conf.d/bulutdemo.ir.key', "'/etc/nginx/conf.d/bulutdemo.key'")
  .replaceAll('ROUTE frontend-bff', 'ROUTE api')
  .replaceAll('ROUTE frontend', 'ROUTE mr-ui')
  .replaceAll('180feedback_frontend_bff_demo', '180feedback_mr_bff_demo')
  .replaceAll('180feedback_frontend_demo', '180feedback_mr_ui_demo')
  .replace('location /bff/ {', 'location /api/ {')
  .replace('proxy_pass                          http://$target:80;', 'proxy_pass                          http://$target:80/;');
const migratedMonorepoNginx = hooks.mergeMonorepoNginxRoutes({
  content: legacyMonorepoNginx,
  serverName: '180feedback.bulutdemo.ir',
  domain: 'bulutdemo.ir',
  projectKey: '180feedback',
  serviceKey: 'front-monorepo',
  environment: 'demo'
});
assert(!migratedMonorepoNginx.includes('ROUTE api'));
assert(!migratedMonorepoNginx.includes('ROUTE mr-ui'));
assert(!migratedMonorepoNginx.includes('180feedback_mr_bff_demo'));
assert(!migratedMonorepoNginx.includes('180feedback_mr_ui_demo'));
assert(migratedMonorepoNginx.includes('ROUTE front-monorepo-bff'));
assert(migratedMonorepoNginx.includes('ROUTE front-monorepo'));
assert(migratedMonorepoNginx.includes('set              $target            180feedback_front_monorepo_bff_demo;'));
assert(migratedMonorepoNginx.includes('set              $target            180feedback_front_monorepo_demo;'));
assert.strictEqual((migratedMonorepoNginx.match(/location \/bff\/ \{/g) || []).length, 1);
assert.strictEqual((migratedMonorepoNginx.match(/location \/api\/ \{/g) || []).length, 0);
assert.strictEqual((migratedMonorepoNginx.match(/location \/ \{/g) || []).length, 1);
assert(!migratedMonorepoNginx.includes('proxy_pass                          http://$target:80/;'));
assert(migratedMonorepoNginx.includes('"/etc/nginx/conf.d/bulutdemo.ir.pem"'));
assert(migratedMonorepoNginx.includes("'/etc/nginx/conf.d/bulutdemo.ir.key'"));
assert(!migratedMonorepoNginx.includes('/etc/nginx/conf.d/bulutdemo.pem'));
assert(!migratedMonorepoNginx.includes('/etc/nginx/conf.d/bulutdemo.key'));
assert.strictEqual(
  hooks.mergeMonorepoNginxRoutes({
    content: migratedMonorepoNginx,
    serverName: '180feedback.bulutdemo.ir',
    domain: 'bulutdemo.ir',
    projectKey: '180feedback',
    serviceKey: 'front-monorepo',
    environment: 'demo'
  }),
  migratedMonorepoNginx
);
const nginxApiSample = hooks.buildNginxSample({
  projectHost: 'locanit',
  projectKey: 'locanit',
  serviceKey: 'api',
  environment: 'dev',
  domain: 'bulutdev.ir'
});
assert(nginxApiSample.includes('server_name locanit.bulutdev.ir;'));
assert(nginxApiSample.includes('location /api/ {'));
assert(nginxApiSample.includes('resolver         127.0.0.11         ipv6=off;'));
assert(nginxApiSample.includes('set              $target            locanit_api_dev;'));
assert(!nginxApiSample.includes('rewrite '));
assert(nginxApiSample.includes('proxy_pass                          http://$target:8080;'));
assert(!nginxApiSample.includes('proxy_pass                          http://$target:8080/;'));
assert(!nginxApiSample.includes('proxy_pass http://locanit_api_dev:8080;'));
assert(nginxApiSample.includes('/etc/nginx/conf.d/bulutdev.ir.pem'));
assert(nginxApiSample.includes('/etc/nginx/conf.d/bulutdev.ir.key'));
assert(!nginxApiSample.includes('/etc/nginx/conf.d/bulutdev.pem'));
assert(nginxApiSample.includes('client_max_body_size 0;'));
assert(nginxApiSample.includes('proxy_set_header Upgrade $http_upgrade;'));
const apiAndMonorepoNginx = hooks.mergeMonorepoNginxRoutes({
  content: nginxApiSample,
  serverName: 'locanit.bulutdev.ir',
  domain: 'bulutdev.ir',
  projectKey: 'locanit',
  serviceKey: 'front-monorepo',
  environment: 'dev'
});
assert(apiAndMonorepoNginx.includes('location /api/ {'));
assert(apiAndMonorepoNginx.includes('set              $target            locanit_api_dev;'));
assert(apiAndMonorepoNginx.includes('location /bff/ {'));
assert(apiAndMonorepoNginx.includes('set              $target            locanit_front_monorepo_bff_dev;'));
assert(apiAndMonorepoNginx.includes('location / {'));
assert(apiAndMonorepoNginx.indexOf('location /api/ {') < apiAndMonorepoNginx.indexOf('location / {'));
assert(apiAndMonorepoNginx.indexOf('location /bff/ {') < apiAndMonorepoNginx.indexOf('location / {'));
const unrelatedBffRoute = nginxApiSample.replace('location /api/ {', 'location /bff/ {');
assert.throws(
  () => hooks.mergeMonorepoNginxRoutes({
    content: unrelatedBffRoute,
    serverName: 'locanit.bulutdev.ir',
    domain: 'bulutdev.ir',
    projectKey: 'locanit',
    serviceKey: 'front-monorepo',
    environment: 'dev'
  }),
  /managed location \/bff\/ belongs to another service/
);
const nginxUiSample = hooks.buildNginxSample({
  projectHost: 'locanit',
  projectKey: 'locanit',
  serviceKey: 'newui',
  environment: 'dev',
  domain: 'bulutdev.ir'
});
assert(nginxUiSample.includes('location / {'));
assert(nginxUiSample.includes('set              $target            locanit_newui_dev;'));
assert(nginxUiSample.includes('proxy_pass                          http://$target:80;'));
assert(!nginxUiSample.includes('proxy_pass                          http://$target:80/;'));
assert(!nginxUiSample.includes('proxy_pass http://locanit_newui_dev:80;'));
const frontRootRoute = hooks.buildNginxRouteBlock({
  projectKey: 'ofe',
  serviceKey: 'front',
  environment: 'dev'
});
assert.strictEqual(frontRootRoute.location, '/');
assert(frontRootRoute.content.includes('set              $target            ofe_front_dev;'));
assert(frontRootRoute.content.includes('proxy_pass                          http://$target:80;'));
assert(!frontRootRoute.content.includes('proxy_pass                          http://$target:80/;'));
const migratedRootSlashSample = hooks.mergeNginxServiceRoute({
  content: nginxUiSample.replace(
    'proxy_pass                          http://$target:80;',
    'proxy_pass                          http://$target:80/;'
  ),
  serverName: 'locanit.bulutdev.ir',
  projectKey: 'locanit',
  serviceKey: 'newui',
  environment: 'dev'
});
assert(migratedRootSlashSample.includes('proxy_pass                          http://$target:80;'));
assert(!migratedRootSlashSample.includes('proxy_pass                          http://$target:80/;'));
const cloudCertificateSample = hooks.buildNginxSample({
  projectHost: 'example',
  projectKey: 'example',
  serviceKey: 'ui',
  environment: 'pro',
  domain: 'bulutco.cloud'
});
const irCertificateSample = hooks.buildNginxSample({
  projectHost: 'example',
  projectKey: 'example',
  serviceKey: 'ui',
  environment: 'pro',
  domain: 'bulutco.ir'
});
assert(cloudCertificateSample.includes('/etc/nginx/conf.d/bulutco.cloud.pem'));
assert(cloudCertificateSample.includes('/etc/nginx/conf.d/bulutco.cloud.key'));
assert(irCertificateSample.includes('/etc/nginx/conf.d/bulutco.ir.pem'));
assert(irCertificateSample.includes('/etc/nginx/conf.d/bulutco.ir.key'));
assert(!cloudCertificateSample.includes('/etc/nginx/conf.d/bulutco.ir.pem'));
const legacyCertificateNginx = [
  nginxApiSample
    .replace('/etc/nginx/conf.d/bulutdev.ir.pem', '"/etc/nginx/conf.d/bulutdev.pem"')
    .replace('/etc/nginx/conf.d/bulutdev.ir.key', "'/etc/nginx/conf.d/bulutdev.key'"),
  'server {',
  '    listen 443 ssl;',
  '    server_name manual.bulutdev.ir;',
  '    ssl_certificate /etc/nginx/conf.d/bulutdev.pem;',
  '    ssl_certificate_key /etc/nginx/conf.d/bulutdev.key;',
  '}',
  ''
].join('\n');
const migratedCertificateNginx = hooks.mergeNginxServiceRoute({
  content: legacyCertificateNginx,
  serverName: 'locanit.bulutdev.ir',
  domain: 'bulutdev.ir',
  projectKey: 'locanit',
  serviceKey: 'api',
  environment: 'dev'
});
assert(migratedCertificateNginx.includes('ssl_certificate "/etc/nginx/conf.d/bulutdev.ir.pem";'));
assert(migratedCertificateNginx.includes("ssl_certificate_key '/etc/nginx/conf.d/bulutdev.ir.key';"));
assert(migratedCertificateNginx.includes('ssl_certificate /etc/nginx/conf.d/bulutdev.pem;'));
assert(migratedCertificateNginx.includes('ssl_certificate_key /etc/nginx/conf.d/bulutdev.key;'));
assert.strictEqual(
  hooks.mergeNginxServiceRoute({
    content: migratedCertificateNginx,
    serverName: 'locanit.bulutdev.ir',
    domain: 'bulutdev.ir',
    projectKey: 'locanit',
    serviceKey: 'api',
    environment: 'dev'
  }),
  migratedCertificateNginx
);
const mergedNginxSample = hooks.mergeNginxServiceRoute({
  content: `${nginxApiSample.replace('    client_max_body_size 0;', '    # manual setting is preserved\n    client_max_body_size 0;')}`,
  serverName: 'locanit.bulutdev.ir',
  projectKey: 'locanit',
  serviceKey: 'newui',
  environment: 'dev'
});
assert(mergedNginxSample.includes('# manual setting is preserved'));
assert(mergedNginxSample.includes('location /api/ {'));
assert(mergedNginxSample.includes('location / {'));
assert(mergedNginxSample.indexOf('location /api/ {') < mergedNginxSample.indexOf('location / {'));
assert.strictEqual((mergedNginxSample.match(/listen 443 ssl;/g) || []).length, 1);
assert.strictEqual((mergedNginxSample.match(/server_name locanit\.bulutdev\.ir;/g) || []).length, 2);
assert.strictEqual(
  hooks.mergeNginxServiceRoute({
    content: mergedNginxSample,
    serverName: 'locanit.bulutdev.ir',
    projectKey: 'locanit',
    serviceKey: 'newui',
    environment: 'dev'
  }),
  mergedNginxSample
);
const rootFirstSample = hooks.mergeNginxServiceRoute({
  content: nginxUiSample,
  serverName: 'locanit.bulutdev.ir',
  projectKey: 'locanit',
  serviceKey: 'api',
  environment: 'dev'
});
assert(rootFirstSample.includes('location /api/ {'));
assert(rootFirstSample.indexOf('location /api/ {') < rootFirstSample.indexOf('location / {'));
const apiRouteBlock = hooks.buildNginxRouteBlock({
  projectKey: 'locanit',
  serviceKey: 'api',
  environment: 'dev'
}).content;
const legacyRootBeforeApiSample = nginxUiSample.replace(
  '    # END PIPELINE-GENERATOR MANAGED ROUTES',
  `${apiRouteBlock}\n    # END PIPELINE-GENERATOR MANAGED ROUTES`
);
const reorderedRootLastSample = hooks.mergeNginxServiceRoute({
  content: legacyRootBeforeApiSample,
  serverName: 'locanit.bulutdev.ir',
  projectKey: 'locanit',
  serviceKey: 'api',
  environment: 'dev'
});
assert(reorderedRootLastSample.indexOf('location /api/ {') < reorderedRootLastSample.indexOf('location / {'));
assert.strictEqual(
  hooks.mergeNginxServiceRoute({
    content: reorderedRootLastSample,
    serverName: 'locanit.bulutdev.ir',
    projectKey: 'locanit',
    serviceKey: 'api',
    environment: 'dev'
  }),
  reorderedRootLastSample
);
const legacyDirectProxySample = mergedNginxSample
  .replace('location /api/ {', 'location /api {')
  .replace(
    /        resolver         127\.0\.0\.11         ipv6=off;\n        set              \$target            locanit_api_dev;\n        proxy_pass                          http:\/\/\$target:8080;/,
    '        proxy_pass http://locanit_api_dev:8080;'
  )
  .replace(
    '        proxy_pass http://locanit_api_dev:8080;',
    '        rewrite          ^/api/(.*)$ /$1 break;\n        proxy_pass http://locanit_api_dev:8080;'
  )
  .replace(
    /        resolver         127\.0\.0\.11         ipv6=off;\n        set              \$target            locanit_newui_dev;\n        proxy_pass                          http:\/\/\$target:80\/?;/,
    '        proxy_pass http://locanit_newui_dev:80;'
  );
const migratedDynamicProxySample = hooks.mergeNginxServiceRoute({
  content: legacyDirectProxySample,
  serverName: 'locanit.bulutdev.ir',
  projectKey: 'locanit',
  serviceKey: 'api',
  environment: 'dev'
});
assert(!migratedDynamicProxySample.includes('proxy_pass http://locanit_api_dev:8080;'));
assert(!migratedDynamicProxySample.includes('proxy_pass http://locanit_newui_dev:80;'));
assert.strictEqual((migratedDynamicProxySample.match(/resolver\s+127\.0\.0\.11\s+ipv6=off;/g) || []).length, 2);
assert(migratedDynamicProxySample.includes('location /api/ {'));
assert(!migratedDynamicProxySample.includes('rewrite '));
assert(migratedDynamicProxySample.includes('proxy_pass                          http://$target:8080;'));
assert(migratedDynamicProxySample.includes('proxy_pass                          http://$target:80;'));
assert(!migratedDynamicProxySample.includes('proxy_pass                          http://$target:80/;'));
assert(migratedDynamicProxySample.indexOf('location /api/ {') < migratedDynamicProxySample.indexOf('location / {'));
assert.strictEqual(
  hooks.mergeNginxServiceRoute({
    content: migratedDynamicProxySample,
    serverName: 'locanit.bulutdev.ir',
    projectKey: 'locanit',
    serviceKey: 'api',
    environment: 'dev'
  }),
  migratedDynamicProxySample
);
const legacySlashNonRootSample = nginxApiSample
  .replace('location /api/ {', 'location /api {')
  .replace(
    'proxy_pass                          http://$target:8080;',
    'rewrite          ^/api/(.*)$ /$1 break;\n        proxy_pass http://$target:8080/;'
  );
const migratedNonRootRewriteSample = hooks.mergeNginxServiceRoute({
  content: legacySlashNonRootSample,
  serverName: 'locanit.bulutdev.ir',
  projectKey: 'locanit',
  serviceKey: 'api',
  environment: 'dev'
});
assert(migratedNonRootRewriteSample.includes('location /api/ {'));
assert(!migratedNonRootRewriteSample.includes('rewrite '));
assert(migratedNonRootRewriteSample.includes('proxy_pass                          http://$target:8080;'));
assert(!migratedNonRootRewriteSample.includes('proxy_pass                          http://$target:8080/;'));
assert.strictEqual(
  hooks.mergeNginxServiceRoute({
    content: migratedNonRootRewriteSample,
    serverName: 'locanit.bulutdev.ir',
    projectKey: 'locanit',
    serviceKey: 'api',
    environment: 'dev'
  }),
  migratedNonRootRewriteSample
);
assert.throws(
  () => hooks.mergeNginxServiceRoute({
    content: `${nginxApiSample}\n${nginxApiSample}`,
    serverName: 'locanit.bulutdev.ir',
    projectKey: 'locanit',
    serviceKey: 'backend',
    environment: 'dev'
  }),
  /multiple HTTPS server blocks/
);
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

  let targetConfigUrl;
  context.fetch = async (url, options = {}) => {
    targetConfigUrl = url;
    assert.strictEqual(options.headers.Authorization, undefined);
    assert.strictEqual(options.headers['X-TFS-FedAuthRedirect'], 'Suppress');
    assert.strictEqual(options.cache, 'no-store');
    assert.strictEqual(options.credentials, 'same-origin');
    assert.strictEqual(options.redirect, 'manual');
    return response({
      body: 'servers:\n  - "QA-192.168.62.153"\nenvironments:\n  - name: qa\n    domain: bulutqa.ir\n',
      url
    });
  };
  const fetchedTargets = await hooks.fetchDeploymentTargets({
    hostUri,
    accessToken: 'extension-session-token'
  });
  assert.strictEqual(
    hooks.buildCollectionUri(hostUri, 'ShonizCollection'),
    'https://azure.example.local/ShonizCollection/'
  );
  assert.strictEqual(
    hooks.buildCollectionUri('https://azure.example.local/tfs/OtherCollection/', 'ShonizCollection'),
    'https://azure.example.local/tfs/ShonizCollection/'
  );
  assert.throws(() => hooks.buildCollectionUri(hostUri, '../unsafe'), /collection name is invalid/);
  assert(targetConfigUrl.startsWith('https://azure.example.local/ShonizCollection/SharedTemplates/'));
  assert(targetConfigUrl.includes('/SharedTemplates/_apis/git/repositories/SharedTemplates/items?'));
  assert(targetConfigUrl.includes('path=%2Fpipeline-generator.yml'));
  assert.deepStrictEqual(Array.from(fetchedTargets.servers), ['QA-192.168.62.153']);
  assert.deepStrictEqual(Array.from(fetchedTargets.environments), ['qa']);

  context.fetch = async (url, options = {}) => {
    assert.strictEqual(options.headers.Authorization, undefined);
    return response({ status: 401, body: 'TF400813: Client authentication required.', url });
  };
  await assert.rejects(
    () => hooks.fetchDeploymentTargets({ hostUri, accessToken: 'must-not-be-forwarded' }),
    (error) =>
      error.status === 401 &&
      error.authenticationMode === 'browser-session' &&
      error.requiredExtensionScope === undefined
  );

  const parsedKomodoCredentials = hooks.parseKomodoCredentialFile(`
KOMODO_ADDRESS=https://komodo.example.local
KOMODO_API_KEY=synthetic-read-key
KOMODO_API_SECRET="synthetic-read-secret"
`);
  assert.strictEqual(parsedKomodoCredentials.address, 'https://komodo.example.local');
  assert.strictEqual(parsedKomodoCredentials.apiKey, 'synthetic-read-key');
  assert.strictEqual(parsedKomodoCredentials.apiSecret, 'synthetic-read-secret');

  const komodoCalls = [];
  context.fetch = async (url, options = {}) => {
    komodoCalls.push({ url, options });
    if (url.includes('path=%2Fkomodo-servers-creds.env')) {
      assert.strictEqual(options.headers.Authorization, undefined);
      assert.strictEqual(options.headers['X-TFS-FedAuthRedirect'], 'Suppress');
      assert.strictEqual(options.credentials, 'same-origin');
      assert.strictEqual(options.redirect, 'manual');
      return response({
        body: [
          'KOMODO_ADDRESS=https://komodo.example.local',
          'KOMODO_API_KEY=synthetic-read-key',
          'KOMODO_API_SECRET=synthetic-read-secret'
        ].join('\n'),
        url
      });
    }
    assert.strictEqual(url, 'https://komodo.example.local/read');
    assert.strictEqual(options.method, 'POST');
    assert.strictEqual(options.headers.Authorization, undefined);
    assert.strictEqual(options.headers['X-Api-Key'], 'synthetic-read-key');
    assert.strictEqual(options.headers['X-Api-Secret'], 'synthetic-read-secret');
    assert.strictEqual(options.credentials, 'omit');
    assert.strictEqual(options.cache, 'no-store');
    assert.strictEqual(JSON.parse(options.body).type, 'ListFullServers');
    return response({
      body: [
        { id: 'server-2', name: 'Production-192.168.0.244', config: { enabled: true } },
        { id: 'server-1', name: 'DEMO-192.168.62.91', config: { enabled: true } },
        { id: 'server-3', name: 'Disabled', config: { enabled: false } },
        { id: 'server-4', name: 'Template', template: true, config: { enabled: true } }
      ],
      url
    });
  };
  const activeKomodoServers = await hooks.fetchKomodoServers({
    hostUri,
    accessToken: 'extension-session-token'
  });
  assert.strictEqual(komodoCalls.length, 2);
  assert(
    komodoCalls[0].url.startsWith(
      'https://azure.example.local/ShonizCollection/SharedTemplates/_apis/git/repositories/SharedTemplates/items?'
    )
  );
  assert.deepStrictEqual(Array.from(activeKomodoServers), [
    'DEMO-192.168.62.91',
    'Production-192.168.0.244'
  ]);

  context.fetch = async (url, options = {}) => {
    if (url.includes('path=%2Fkomodo-servers-creds.env')) {
      return response({
        body: [
          'KOMODO_ADDRESS=https://komodo.example.local',
          'KOMODO_API_KEY=synthetic-read-key',
          'KOMODO_API_SECRET=synthetic-read-secret'
        ].join('\n'),
        url
      });
    }
    const request = JSON.parse(options.body);
    assert.strictEqual(request.type, 'ListDockerNetworks');
    assert.strictEqual(request.params.server, 'DEMO-192.168.62.91');
    assert.strictEqual(options.headers['X-Api-Key'], 'synthetic-read-key');
    assert.strictEqual(options.headers['X-Api-Secret'], 'synthetic-read-secret');
    return response({
      body: [
        { name: 'bridge', driver: 'bridge' },
        { name: 'nginx-net', driver: 'bridge' }
      ],
      url
    });
  };
  assert.strictEqual(
    await hooks.resolveNginxNetworkForServer({
      hostUri,
      server: 'DEMO-192.168.62.91'
    }),
    'nginx-net'
  );

  environment.options = [];
  elements.get('komodoServer').options = [];
  context.fetch = async (url) => {
    if (url.includes('path=%2Fkomodo-servers-creds.env')) {
      return response({
        body: [
          'KOMODO_ADDRESS=https://komodo.example.local',
          'KOMODO_API_KEY=synthetic-read-key',
          'KOMODO_API_SECRET=synthetic-read-secret'
        ].join('\n'),
        url
      });
    }
    if (url === 'https://komodo.example.local/read') {
      return response({
        body: [
          { id: 'qa-id', name: 'QA-192.168.62.153', config: { enabled: true } },
          { id: 'demo-id', name: 'DEMO-192.168.62.91', config: { enabled: true } }
        ],
        url
      });
    }
    return response({
      body: 'environments:\n  - name: demo\n    domain: bulutdemo.ir\n  - name: qa\n    domain: bulutqa.ir\n',
      url
    });
  };
  await hooks.loadDeploymentTargets({
    hostUri,
    accessToken: 'extension-session-token',
    branch: 'feature/qa'
  });
  assert.strictEqual(environment.value, 'qa');
  assert.strictEqual(elements.get('komodoServer').value, 'QA-192.168.62.153');
  assert.strictEqual(hooks.state.deploymentTargetsReady, true);

  const supportRepos = new Map();
  const supportPushes = new Map();
  context.fetch = async (url, options = {}) => {
    const method = options.method || 'GET';
    if (url.includes('/_apis/git/repositories?') && method === 'GET') {
      return response({ body: { value: Array.from(supportRepos.values()) }, url });
    }
    if (url.includes('/_apis/git/repositories?') && method === 'POST') {
      const body = JSON.parse(options.body);
      const id = body.name.includes('_Docker_') ? 'docker-repo-id' : 'nginx-repo-id';
      const created = { id, name: body.name };
      supportRepos.set(body.name, created);
      return response({ body: created, url });
    }
    if (url.includes('/refs?') && method === 'GET') {
      return response({ body: { value: [] }, url });
    }
    if (url.includes('/pushes?') && method === 'POST') {
      const repoId = url.includes('docker-repo-id') ? 'docker-repo-id' : 'nginx-repo-id';
      supportPushes.set(repoId, JSON.parse(options.body));
      return response({ body: { pushId: supportPushes.size }, url });
    }
    if (method === 'PATCH' && /\/repositories\/(docker|nginx)-repo-id\?/.test(url)) {
      assert.strictEqual(JSON.parse(options.body).defaultBranch, 'refs/heads/main');
      return response({ body: { id: url.includes('docker-repo-id') ? 'docker-repo-id' : 'nginx-repo-id' }, url });
    }
    throw new Error(`Unexpected support repository request: ${method} ${url}`);
  };
  const supportResults = await hooks.ensureSupportRepositories({
    hostUri,
    projectId,
    projectName: 'RideSharing',
    environment: 'demo',
    domain: 'bulutdemo.ir',
    service: 'api',
    repositoryAddress: 'registry.buluttakin.com',
    accessToken: 'extension-session-token'
  });
  assert.deepStrictEqual(
    Array.from(supportResults, (result) => result.repo.name),
    ['RideSharing_Docker_DevOps', 'RideSharing_Nginx_DevOps']
  );
  const dockerPaths = supportPushes
    .get('docker-repo-id')
    .commits[0].changes.map((change) => change.item.path);
  const nginxPaths = supportPushes
    .get('nginx-repo-id')
    .commits[0].changes.map((change) => change.item.path);
  assert.deepStrictEqual(dockerPaths, ['/demo_ridesharing/compose.yml']);
  assert.deepStrictEqual(nginxPaths, ['/demo/ridesharing-demo.conf']);
  const dockerComposeContent = supportPushes
    .get('docker-repo-id')
    .commits[0].changes.find((change) => change.item.path.endsWith('/compose.yml')).newContent.content;
  assert(dockerComposeContent.includes('container_name: ridesharing_api_demo'));
  assert(dockerComposeContent.includes('registry.buluttakin.com/ridesharing/api-demo:${IMAGE_TAG:-CHANGE_ME}'));
  const nginxContent = supportPushes
    .get('nginx-repo-id')
    .commits[0].changes.find((change) => change.item.path.endsWith('.conf')).newContent.content;
  assert(nginxContent.includes('server_name ridesharing.bulutdemo.ir;'));
  assert(nginxContent.includes('location /api/ {'));
  assert(nginxContent.includes('set              $target            ridesharing_api_demo;'));
  assert(!nginxContent.includes('rewrite '));
  assert(nginxContent.includes('proxy_pass                          http://$target:8080;'));
  assert(!nginxContent.includes('proxy_pass http://ridesharing_api_demo:8080;'));
  hooks.state.rawProjectName = 'RideSharing';
  hooks.state.projectName = 'RideSharing';
  hooks.showCompletionLinks({
    supportRepositories: supportResults,
    pipelineDefinition: { id: 344, name: filename }
  });
  assert(elements.get('nginx-result-link').href.includes('/RideSharing/_git/RideSharing_Nginx_DevOps?'));
  assert(elements.get('nginx-result-link').href.includes('path=%2Fdemo%2Fridesharing-demo.conf'));
  assert(elements.get('compose-result-link').href.includes('path=%2Fdemo_ridesharing%2Fcompose.yml'));
  assert.strictEqual(
    elements.get('pipeline-result-link').href,
    `${hostUri}RideSharing/_build?definitionId=344`
  );
  assert.strictEqual(hooks.state.provisioningComplete, true);
  assert.strictEqual(form.hidden, true);
  assert.strictEqual(submitButton.disabled, true);

  const existingSupportCalls = [];
  context.fetch = async (url, options = {}) => {
    const method = options.method || 'GET';
    existingSupportCalls.push({ url, method });
    if (url.includes('/refs?')) {
      return response({ body: { value: [{ objectId: '2222222222222222222222222222222222222222' }] }, url });
    }
    if (url.includes('/items?') && url.includes('%24format=text')) {
      return response({ body: 'mattermost_channel=already-configured', url });
    }
    throw new Error(`Existing support repository content must not be written: ${method} ${url}`);
  };
  const existingBootstrap = await hooks.ensureRepositoryBootstrapFiles({
    hostUri,
    projectId,
    repo: { id: 'docker-repo-id', name: 'RideSharing_Docker_DevOps' },
    directory: 'demo_ridesharing',
    sampleFile: {
      path: '/demo_ridesharing/compose.yml',
      content: 'services: {}\n'
    },
    accessToken: 'extension-session-token'
  });
  assert.strictEqual(existingBootstrap.skipped, true);
  assert(existingSupportCalls.every(({ method }) => method === 'GET'));

  let nginxMergePush;
  context.fetch = async (url, options = {}) => {
    const method = options.method || 'GET';
    if (url.includes('/refs?') && method === 'GET') {
      return response({ body: { value: [{ objectId: '3333333333333333333333333333333333333333' }] }, url });
    }
    if (url.includes('/items?') && method === 'GET') {
      const filePath = new URL(url).searchParams.get('path');
      return response({
        body: nginxApiSample,
        url
      });
    }
    if (url.includes('/pushes?') && method === 'POST') {
      nginxMergePush = JSON.parse(options.body);
      return response({ body: { pushId: 3 }, url });
    }
    throw new Error(`Unexpected Nginx merge request: ${method} ${url}`);
  };
  const mergedBootstrap = await hooks.ensureRepositoryBootstrapFiles({
    hostUri,
    projectId,
    repo: { id: 'nginx-repo-id', name: 'Locanit_Nginx_DevOps' },
    directory: 'dev',
    sampleFile: {
      path: '/dev/locanit-dev.conf',
      content: nginxUiSample,
      mergeExisting: (content) => hooks.mergeNginxServiceRoute({
        content,
        serverName: 'locanit.bulutdev.ir',
        projectKey: 'locanit',
        serviceKey: 'newui',
        environment: 'dev'
      })
    },
    accessToken: 'extension-session-token'
  });
  assert.strictEqual(mergedBootstrap.skipped, false);
  assert.strictEqual(nginxMergePush.commits[0].changes.length, 1);
  assert.strictEqual(nginxMergePush.commits[0].changes[0].changeType, 'edit');
  assert.strictEqual(nginxMergePush.commits[0].changes[0].item.path, '/dev/locanit-dev.conf');
  assert(nginxMergePush.commits[0].changes[0].newContent.content.includes('location / {'));

  let monorepoGeneratedPush;
  context.fetch = async (url, options = {}) => {
    const method = options.method || 'GET';
    if (url.includes('/refs?') && method === 'GET') {
      return response({ body: { value: [{ objectId: '4444444444444444444444444444444444444444' }] }, url });
    }
    if (url.includes('/items?') && method === 'GET') {
      const filePath = new URL(url).searchParams.get('path');
      if (filePath === '/.devops/deployments.yml') {
        return response({ body: 'version: 1\n# operator customization\n', url });
      }
      return response({ body: 'old pipeline yaml\n', url });
    }
    if (url.includes('/pushes?') && method === 'POST') {
      monorepoGeneratedPush = JSON.parse(options.body);
      return response({ body: { pushId: 4 }, url });
    }
    throw new Error(`Unexpected generated Monorepo request: ${method} ${url}`);
  };
  const generatedFiles = await hooks.postGeneratedFiles({
    hostUri,
    projectId,
    repoId: repo.id,
    accessToken: 'extension-session-token',
    files: [
      { path: '/mr.yml', content: 'new pipeline yaml\n' },
      { path: '/.devops/deployments.yml', content: monorepoContract, overwrite: false }
    ]
  });
  assert.strictEqual(generatedFiles.skipped, false);
  assert.deepStrictEqual(
    monorepoGeneratedPush.commits[0].changes.map((change) => [change.changeType, change.item.path]),
    [['edit', '/mr.yml']]
  );
  assert(
    !monorepoGeneratedPush.commits[0].changes.some(
      (change) => change.item.path === '/.devops/deployments.yml'
    ),
    'Operator-edited deployments.yml must never be overwritten.'
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

  const serviceLessMigrationCalls = [];
  context.fetch = async (url, options = {}) => {
    const method = options.method || 'GET';
    serviceLessMigrationCalls.push({ url, method, options });
    if (url.includes('/_apis/pipelines?')) {
      return response({ body: { value: [{ id: 346, name: serviceLessFilename }] }, url });
    }
    if (url.includes('/_apis/build/definitions/346') && method === 'GET') {
      return response({
        body: {
          id: 346,
          revision: 3,
          name: serviceLessFilename,
          path: '\\KOMODO',
          process: { type: 2, yamlFilename: `/${serviceLessFilename}` },
          repository: { id: repo.id, name: repo.name, type: 'TfsGit', defaultBranch: 'refs/heads/main' }
        },
        url
      });
    }
    if (url.includes('/_apis/build/definitions/346') && method === 'PUT') {
      const body = JSON.parse(options.body);
      assert.strictEqual(body.id, 346);
      assert.strictEqual(body.revision, 3);
      assert.strictEqual(body.name, filename);
      assert.strictEqual(body.process.yamlFilename, `/${filename}`);
      return response({ body: { ...body, id: 346, revision: 4 }, url });
    }
    throw new Error(`Unexpected Service-less migration request: ${method} ${url}`);
  };
  const serviceAwareMigration = await hooks.upsertPipelineDefinition({
    hostUri,
    projectId,
    repo,
    pipelineName: filename,
    pipelinePath: `/${filename}`,
    legacyPipelineNames: [serviceLessFilename],
    legacyPipelinePaths: [`/${serviceLessFilename}`],
    branch: 'main',
    accessToken: 'test-token'
  });
  assert.strictEqual(serviceAwareMigration.id, 346);
  assert(
    serviceLessMigrationCalls.some(
      ({ url, method }) => url.includes('/_apis/build/definitions/346') && method === 'PUT'
    )
  );

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
              process: { type: 2, yamlFilename: `/${legacyFilename}` },
              repository: { id: repo.id, defaultBranch: 'refs/heads/main' }
            },
            {
              id: 344,
              revision: 7,
              name: 'RideSharing_RideSharing_Backend_demo',
              path: '\\KOMODO',
              process: { type: 2, yamlFilename: `/${legacyFilename}` },
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
          process: { type: 2, yamlFilename: `/${legacyFilename}` },
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
    legacyPipelineNames: [serviceLessFilename, previousEnvironmentFirstFilename, legacyFilename],
    legacyPipelinePaths: [
      `/${serviceLessFilename}`,
      `/${previousEnvironmentFirstFilename}`,
      `/${legacyFilename}`
    ],
    branch: 'main',
    accessToken: 'test-token'
  });
  assert.strictEqual(migrated.id, 344);
  const migrationYamlLookups = migrationCalls
    .filter(({ url, method }) => url.includes('/_apis/build/definitions?') && method === 'GET')
    .map(({ url }) => new URL(url).searchParams.get('yamlFilename'));
  assert.deepStrictEqual(migrationYamlLookups, [
    `/${filename}`,
    `/${serviceLessFilename}`,
    `/${previousEnvironmentFirstFilename}`,
    `/${legacyFilename}`
  ]);
  assert(migrationCalls.some(({ url, method }) => url.includes('/_apis/build/definitions/344') && method === 'PUT'));
  assert(!migrationCalls.some(({ url, method }) => url.includes('/_apis/pipelines/344') && method === 'PUT'));

  const releaseCalls = [];
  let updatedReleaseBody;
  context.fetch = async (url, options = {}) => {
    const method = options.method || 'GET';
    const parsed = new URL(url);
    releaseCalls.push({ url, method, options });
    if (url.includes('/_apis/release/definitions?') && method === 'GET' && parsed.searchParams.has('searchText')) {
      assert.strictEqual(parsed.searchParams.get('searchText'), 'API DEMO');
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
      assert.strictEqual(body.name, 'API DEMO');
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
    service: 'api',
    environment: 'demo',
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
      assert.strictEqual(parsed.searchParams.get('searchText'), 'API DEMO');
      return response({ body: { value: [{ id: 5, name: 'API DEMO' }] }, url });
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
    service: 'api',
    environment: 'demo',
    branch: 'main',
    queueName: 'PublishDockerAgent',
    accessToken: 'test-token'
  });
  assert.strictEqual(reusedRelease.id, 5);
  assert.strictEqual(reusedRelease.created, false);
  assert.strictEqual(reusedRelease.updated, false);
  assert(noOpReleaseCalls.every(({ method }) => method === 'GET'));

console.log(
  'UI behavior regression tests passed: Environment/domain parsing, direct enabled-server discovery, Service-aware BranchToEnvironment Pipeline naming with legacy migration, root-last Nginx routing with managed rewrite removal, idempotent Compose/shared-route merging, locked completion links, Service-aware Release naming, and Pipeline/Release/KomodoAPI reconciliation.'
);
};

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
