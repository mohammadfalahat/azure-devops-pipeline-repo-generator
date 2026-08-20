#!/usr/bin/env node

/*
 * Offline validation for the static extension assets. It deliberately does not
 * contact Azure DevOps; REST/API behavior must be verified against the target
 * Azure DevOps Server after publishing the VSIX.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const yaml = require('js-yaml');

const root = path.resolve(__dirname, '..');
const read = (relativePath) => fs.readFileSync(path.join(root, relativePath), 'utf8');
const fail = (message) => {
  throw new Error(`[extension validation] ${message}`);
};

const manifest = JSON.parse(read('vss-extension.json'));
const readme = read('README.md');
const developerDocs = [
  'docs/monorepo-service-deployment-fa.md',
  'docs/architecture.md',
  'docs/rest-api-contracts.md',
  'docs/development-and-operations.md'
];

const centralMonorepoTemplatePath = 'examples/shared-templates/monorepo/pipeline.yml';
if (!fs.existsSync(path.join(root, centralMonorepoTemplatePath))) {
  fail(`central Monorepo template example is missing: ${centralMonorepoTemplatePath}`);
}
const centralMonorepoNginxPath = 'examples/shared-templates/monorepo/nginx/default.conf';
if (!fs.existsSync(path.join(root, centralMonorepoNginxPath))) {
  fail(`central Monorepo Nginx example is missing: ${centralMonorepoNginxPath}`);
}
const centralMonorepoTemplate = read(centralMonorepoTemplatePath);
let centralMonorepoTemplateDocument;
try {
  centralMonorepoTemplateDocument = yaml.load(centralMonorepoTemplate);
} catch (error) {
  fail(`central Monorepo template is not valid YAML: ${error.message}`);
}
for (const requiredTemplateBehavior of [
  'monorepo/mr-build.cjs',
  'monorepo/nginx/default.conf',
  'mr-drop/runtime/nginx',
  'ListFullRepos',
  'CreateRepo',
  'UpdateRepo',
  'ListFullStacks',
  'CreateStack',
  'UpdateStack',
  'linked_repo:$linked_repo',
  'server_id:$server,path:$path',
  'files_on_host:false',
  'MR_KOMODO_STACK',
  'PublishBuildArtifacts@1'
]) {
  if (!centralMonorepoTemplate.includes(requiredTemplateBehavior)) {
    fail(`central Monorepo template is missing ${requiredTemplateBehavior}.`);
  }
}
if (centralMonorepoTemplate.includes('"DeployStack"')) {
  fail('central Monorepo Build template must configure GitOps resources without deploying the Stack.');
}
const ensureKomodoStep = centralMonorepoTemplateDocument?.stages?.[0]?.jobs?.[0]?.steps?.find(
  (step) => step.displayName === 'Ensure Komodo GitOps repository and shared Docker Stack'
);
if (!ensureKomodoStep?.inputs?.script) {
  fail('central Monorepo template must contain the Komodo GitOps upsert Bash step.');
}

for (const docPath of developerDocs) {
  if (!fs.existsSync(path.join(root, docPath))) {
    fail(`required developer documentation is missing: ${docPath}`);
  }
  if (!readme.includes(`](${docPath})`)) {
    fail(`README must link to ${docPath}.`);
  }
}

if (!readme.includes(`**${manifest.version}**`)) {
  fail(`README must identify the current manifest version ${manifest.version}.`);
}

const requiredScopes = [
  'vso.code',
  'vso.code_manage',
  'vso.project',
  'vso.build',
  'vso.build_execute',
  'vso.release',
  'vso.release_manage',
  'vso.agentpools',
  'vso.serviceendpoint',
  'vso.variablegroups_read'
];
for (const scope of requiredScopes) {
  if (!manifest.scopes?.includes(scope)) {
    fail(`manifest is missing required scope: ${scope}`);
  }
}

const dialogContribution = manifest.contributions?.find(
  (contribution) => contribution.id === 'pipeline-generator-dialog'
);
if (
  dialogContribution?.type !== 'ms.vss-web.control' ||
  dialogContribution?.properties?.uri !== 'dist/index.html'
) {
  fail('manifest must declare dist/index.html as the pipeline-generator-dialog host control.');
}

const hubContribution = manifest.contributions?.find(
  (contribution) => contribution.id === 'pipeline-generator-hub'
);
if (
  hubContribution?.type !== 'ms.vss-web.hub' ||
  hubContribution?.properties?.uri !== 'dist/index.html' ||
  !hubContribution?.targets?.includes('ms.vss-code-web.code-hub-group')
) {
  fail('manifest must declare dist/index.html as an Azure Repos Pipeline Generator hub.');
}

const monorepoAction = manifest.contributions?.find(
  (contribution) => contribution.id === 'generate-monorepo-action'
);
if (
  monorepoAction?.type !== 'ms.vss-web.action' ||
  monorepoAction?.properties?.registeredObjectId !== 'generate-monorepo-action' ||
  monorepoAction?.properties?.text !== 'Generate MonoRepo'
) {
  fail('manifest must declare the separate Generate MonoRepo branch action.');
}

const menuAction = read('dist/menu-action.js');
for (const requiredDialogImplementation of [
  'ms.vss-features.host-page-layout-service',
  'ms.vss-features.host-navigation-service',
  'buildDialogContributionId',
  'buildGeneratorHubUrl',
  'openCustomDialog',
  'configuration: bootstrapPayload',
  '/_apps/hub/'
]) {
  if (!menuAction.includes(requiredDialogImplementation)) {
    fail(`menu-action.js is missing host-dialog behavior: ${requiredDialogImplementation}.`);
  }
}
if (/window\.open\s*\(|openWindow\s*\(|getAccessTokenWithRetry|postBootstrapMessage/.test(menuAction)) {
  fail('the branch action must never open a detached generator or acquire/transfer a token.');
}
if (
  !menuAction.includes("sdk.register('generate-monorepo-action'") ||
  !menuAction.includes("params.set('mode', mode)") ||
  !menuAction.includes("buildAction('monorepo')")
) {
  fail('menu-action.js must register and propagate the separate monorepo mode.');
}

const releaseConfigContext = { window: {} };
vm.runInNewContext(read('dist/release-config.js'), releaseConfigContext, { filename: 'dist/release-config.js' });
const releaseConfig = releaseConfigContext.window.PipelineGeneratorReleaseConfig;
if (!releaseConfig?.enabled || releaseConfig.folder !== '\\komodo') {
  fail('release configuration must enable the \\komodo release folder.');
}

if (Object.prototype.hasOwnProperty.call(releaseConfig, 'nameSuffix')) {
  fail('Release names must come from uppercased service and environment, not a Pipeline-name suffix.');
}
if (releaseConfig.variableGroupName !== 'KomodoAPI') {
  fail('release configuration must link the KomodoAPI variable group.');
}
const requiredReleaseVariables = ['AZP_TOKEN', 'KOMODO_API_KEY', 'KOMODO_API_SECRET'];
for (const variableName of requiredReleaseVariables) {
  if (!releaseConfig.requiredVariableNames?.includes(variableName)) {
    fail(`release configuration is missing required variable name: ${variableName}.`);
  }
}
const scriptSource = releaseConfig.scriptSource || {};
if (scriptSource.type === 'inline') {
  if (!scriptSource.content?.trim()) {
    fail('release configuration must provide a non-empty inline Bash script.');
  }
} else if (scriptSource.type === 'packagedFile') {
  if (!String(scriptSource.path || '').trim()) {
    fail('packagedFile script source requires path.');
  }
} else if (scriptSource.type === 'azureReposFile') {
  for (const field of ['project', 'repository', 'path']) {
    if (!String(scriptSource[field] || '').trim()) {
      fail(`azureReposFile script source requires ${field}.`);
    }
  }
} else {
  fail('release scriptSource.type must be inline, packagedFile, or azureReposFile.');
}

const packagedReleaseScript = read('dist/release-inline-task.sh');
const documentedReleaseScript = read('scripts/release-inline-task.example.sh');
if (packagedReleaseScript.trimEnd() !== documentedReleaseScript.trimEnd()) {
  fail('packaged and documented Release inline Bash scripts must be identical.');
}
for (const variableName of requiredReleaseVariables) {
  if (!packagedReleaseScript.includes(`\$(${variableName})`)) {
    fail(`packaged Release Bash script must reference Azure DevOps variable macro $(${variableName}).`);
  }
}
if (!packagedReleaseScript.includes('SCRIPT_PATH="release-komodo.sh"')) {
  fail('packaged Release Bash script must execute release-komodo.sh from SharedTemplates.');
}
const authenticatedGitIndex = packagedReleaseScript.indexOf('GIT_CONFIG_VALUE_0="Authorization: Basic ${AUTH_HDR}"');
const firstXtraceIndex = packagedReleaseScript.indexOf('set -x');
if (authenticatedGitIndex < 0 || (firstXtraceIndex >= 0 && firstXtraceIndex < authenticatedGitIndex)) {
  fail('secret-derived Git Authorization header must never be emitted under shell xtrace.');
}
if (
  packagedReleaseScript.includes('git -c http.extraHeader=') ||
  /curl[^\n]*-H\s+["']Authorization: Basic/.test(packagedReleaseScript) ||
  /wget[^\n]*--header=/.test(packagedReleaseScript)
) {
  fail('secret-derived Authorization headers must not be placed in process arguments.');
}

const html = read('dist/index.html');
if (!/release-config\.js[\s\S]*ui\.js/.test(html)) {
  fail('index.html must load release-config.js before ui.js.');
}
if (
  !html.includes('id="authorize-extension"') ||
  !html.includes('Open extension authorization') ||
  !html.includes('id="reauthenticate"') ||
  !html.includes('Sign out and authenticate again')
) {
  fail('index.html must provide extension authorization and Azure DevOps session restart actions.');
}
if (/id="pat-token"|Use PAT/i.test(html)) {
  fail('the browser UI must not request a Personal Access Token.');
}
if (
  !html.includes('Loading from pipeline-generator.yml') ||
  !html.includes('Loading active Komodo servers') ||
  /<option value="(?:dev|demo|qa|pro|DEMO-192|Development-192|Production-)/.test(html)
) {
  fail('Environment and active Komodo server choices must use dynamic loading placeholders.');
}
if (
  !html.includes('id="completion-panel"') ||
  !html.includes('tabindex="-1"') ||
  !html.includes('id="nginx-result-link"') ||
  !html.includes('id="compose-result-link"') ||
  !html.includes('id="pipeline-result-link"') ||
  !html.includes('id="contract-result-link"')
) {
  fail('index.html must expose review links for Nginx, Compose, MR contract, and Pipeline results.');
}

const monorepoBuildScript = read('dist/monorepo-build.cjs');
for (const requiredBuildBehavior of [
  'pnpm install --frozen-lockfile',
  'show projects --affected',
  'deploy:shell',
  'deploy:bff',
  'continue_on_module_error',
  'successfulProjects',
  'failedProjects',
  'task.complete result=SucceededWithIssues',
  'bffProject',
  'komodoStack',
  'staticContainer',
  'manifest.json',
  'inventory.tsv',
  'modules.tsv'
]) {
  if (!monorepoBuildScript.includes(requiredBuildBehavior)) {
    fail(`Monorepo build script is missing ${requiredBuildBehavior}.`);
  }
}

const monorepoReleaseScript = read('dist/monorepo-release-inline-task.sh');
for (const requiredReleaseBehavior of [
  '$(AZP_TOKEN)',
  '$(KOMODO_API_KEY)',
  '$(KOMODO_API_SECRET)',
  '/terminal/execute',
  'artifactName=mr-drop',
  'composeRepository',
  'komodoStack',
  'bffProject',
  'DeployStack',
  'GetUpdate',
  '.success == true',
  'Komodo Stack deployed from ADO Git',
  'current symlink and runtime Nginx configuration rolled back',
  'MR orphan retained for manual review',
  'ln -s "../modules/$module_name"',
  'runtime/nginx/default.conf',
  'nginx-default.conf',
  'nginx:1.27-alpine -t',
  'nginx -s reload',
  'mv -Tf "$link" "$root/current"',
  'server: process.env.MR_TERMINAL_SERVER'
]) {
  if (!monorepoReleaseScript.includes(requiredReleaseBehavior)) {
    fail(`Monorepo Release script is missing ${requiredReleaseBehavior}.`);
  }
}
if (monorepoReleaseScript.includes('target: { type: \'Server\'')) {
  fail('Monorepo Release must use the Komodo 1.19.5 terminal payload, not the v2 target/init shape.');
}
if (/docker compose\s/.test(monorepoReleaseScript)) {
  fail('Monorepo Release must deploy the ADO Git-managed Compose through Komodo DeployStack.');
}
if (/set\s+-x/.test(monorepoReleaseScript) || /curl[^\n]*-H\s+/.test(monorepoReleaseScript)) {
  fail('Monorepo Release must not expose secret headers through xtrace or process arguments.');
}

for (const removedProxyAsset of [
  'scripts/komodo-server-proxy.js',
  'dist/komodo-server-config.js',
  'deploy/komodo-server-proxy/nginx-location.conf'
]) {
  if (fs.existsSync(path.join(root, removedProxyAsset))) {
    fail(`obsolete Komodo proxy asset must be removed: ${removedProxyAsset}`);
  }
}

const ui = read('dist/ui.js');
for (const requiredImplementation of [
  'ensureReleaseDefinition',
  'buildMonorepoPipelineYaml',
  'buildMonorepoKomodoResourceNames',
  'buildMonorepoDeploymentContract',
  'buildMonorepoSupportRepositorySpecs',
  'buildMonorepoComposeSample',
  'buildMonorepoNginxSample',
  'mergeMonorepoNginxRoutes',
  'postGeneratedFiles',
  'monorepo-release-inline-task.sh',
  'parseDeploymentTargetsYaml',
  'fetchDeploymentTargets',
  'buildCollectionUri',
  'buildCentralGitItemUrl',
  'centralGitRequestOptions',
  "CENTRAL_COLLECTION_NAME = 'ShonizCollection'",
  'parseKomodoCredentialFile',
  'fetchKomodoCredentials',
  'extractEnabledKomodoServers',
  'fetchKomodoServers',
  'KOMODO_CREDENTIAL_CONFIG',
  '/komodo-servers-creds.env',
  'ListFullServers',
  'record.config?.enabled === true',
  'DEPLOYMENT_TARGETS_CONFIG',
  'buildSupportRepositorySpecs',
  'buildComposeSample',
  'tokenizeNginx',
  'findNginxHttpsServer',
  'mergeNginxServiceRoute',
  'buildNginxSample',
  'normalizeNginxManagedRoutes',
  'findManagedRootRouteIndex',
  "const location = requestedLocation || (frontend ? '/' : `/${serviceKey}/`);",
  'resolver         127.0.0.11         ipv6=off;',
  'proxy_pass                          http://$target:${internalPort}/;',
  'proxy_pass                          http://$target:${internalPort};',
  'generatedRewritePattern',
  'ensureSupportRepositories',
  'ensureRepositoryBootstrapFiles',
  'showCompletionLinks',
  'finishProvisioning',
  'state.provisioningComplete',
  'form.hidden = true',
  'buildPipelineFilename',
  'buildLegacyPipelineFilename',
  'buildLegacyEnvironmentFirstPipelineFilename',
  'Environment is required to build the Pipeline filename.',
  'legacyPipelinePaths',
  'buildReleaseName',
  'buildReleaseDefinitionPayload',
  'updateBuildDefinition',
  'getBuildDefinitionByYamlPath',
  'updateReleaseDefinition',
  'getReleaseDefinitionByPipelineId',
  'resolveReleaseVariableGroup',
  'vso.variablegroups_read',
  'BASH_TASK_ID',
  'runProvisioningStep',
  'buildSignOutUrl',
  'buildExtensionManagementUrl',
  'getDialogConfiguration',
  'getHostNavigationState',
  'navigateHost',
  'openExtensionAuthorization',
  'isHostAuthorizationError',
  'buildTokenRecoveryMessage',
  'markRequiredExtensionScope',
  'restartAzureDevOpsSession'
]) {
  if (!ui.includes(requiredImplementation)) {
    fail(`ui.js is missing ${requiredImplementation}.`);
  }
}
for (const centralTemplateReference of [
  'repository: SharedTemplatesRepo',
  'endpoint: ShonizCollection',
  'template: monorepo/pipeline.yml@SharedTemplatesRepo'
]) {
  if (!ui.includes(centralTemplateReference)) {
    fail(`Monorepo Pipeline renderer is missing central template reference: ${centralTemplateReference}`);
  }
}
if (ui.includes("{ path: '/.devops/mr-build.cjs'")) {
  fail('The extension must not copy mr-build.cjs into project repositories after centralization.');
}
if (ui.includes('proxy_pass http://${containerName}')) {
  fail('managed Nginx routes must never proxy directly to a compiled container hostname.');
}
if (/localStorage|sessionStorage|document\.cookie/.test(ui)) {
  fail('ui.js must not persist credentials in browser storage or cookies.');
}
if (/["'](?:K|S)-[A-Za-z0-9]{20,}["']/.test(ui)) {
  fail('ui.js must not contain a literal Komodo API credential.');
}
if (ui.includes("scheme: 'basic-pat'") || ui.includes('createPatCredential') || ui.includes('connectWithPat')) {
  fail('the browser UI must use only the Azure DevOps host token, never a PAT fallback.');
}
if (!ui.includes("_signout") || !ui.includes('state.accessToken = null')) {
  fail('session restart must discard the in-memory host token and navigate to Azure DevOps sign-out.');
}
if (!ui.includes('_settings/extensions?tab=installed')) {
  fail('hosted UI must link to collection-level extension authorization.');
}
if (/\b(?:sdk|VSS)\.resize\s*\(/.test(ui)) {
  fail('Hub UI must not request host resizing; legacy VSS.resize creates a width feedback loop on the target Server.');
}
const styles = read('dist/styles.css');
if (!/html\s*\{[\s\S]*?height:\s*100%[\s\S]*?overflow:\s*hidden/.test(styles)) {
  fail('the Hub document must stay bounded to the host viewport.');
}
if (!/body\s*\{[\s\S]*?height:\s*100%[\s\S]*?overflow:\s*hidden/.test(styles)) {
  fail('the Hub body must remain bounded instead of relying on iframe root scrolling.');
}
if (
  !/\.wrapper\s*\{[\s\S]*?position:\s*fixed[\s\S]*?inset:\s*0[\s\S]*?height:\s*100%[\s\S]*?overflow-y:\s*scroll/.test(
    styles
  )
) {
  fail('the complete Hub form must use an explicit full-viewport vertical scroll container.');
}
if (/setTimeout\s*\([\s\S]{0,300}navigateHost\([\s\S]{0,200}_build\?definitionId/.test(ui)) {
  fail('successful provisioning must show review links and must not redirect automatically to Pipelines.');
}
if (!ui.includes('const hasHostContext = isFramed;')) {
  fail('host-dialog SDK initialization must not depend on document.referrer being present.');
}

if (!ui.includes("const PIPELINE_API_VERSION = '7.1-preview.1'")) {
  fail('ui.js must use the Azure DevOps Server-compatible preview Pipelines API contract.');
}
if (!ui.includes("const BUILD_API_VERSION = '7.1'")) {
  fail('ui.js must use Build Definitions API 7.1 for Azure DevOps Server-compatible Pipeline updates.');
}
if (!ui.includes("searchParams.set('repositoryId', repositoryId)")) {
  fail('pipeline creation must send repositoryId so Azure DevOps Server can bind the generated YAML repository.');
}
if (!ui.includes("repositoryId: repo.id")) {
  fail('pipeline creation must pass the generated repository ID to the Pipelines API URL builder.');
}
if (!ui.includes('readPipelineResponse')) {
  fail('pipeline create responses must be validated for a pipeline ID.');
}
if (
  !ui.includes('const buildPipelineName = (pipelineFilename) => pipelineFilename;') ||
  !ui.includes('const pipelineName = buildPipelineName(pipelineFilename);')
) {
  fail('Pipeline display name must exactly match the generated YAML filename.');
}
if (
  !ui.includes("project: 'SharedTemplates'") ||
  !ui.includes("repository: 'SharedTemplates'") ||
  !ui.includes("path: '/pipeline-generator.yml'") ||
  !ui.includes("branch: 'main'")
) {
  fail('Deployment target configuration must come from SharedTemplates/SharedTemplates:/pipeline-generator.yml@main.');
}
if (
  !ui.includes("`${compactProject}_Docker_DevOps`") ||
  !ui.includes("`${compactProject}_Nginx_DevOps`") ||
  ui.includes("{ path: '/environments', content: 'mattermost_channel=changeme' }") ||
  !ui.includes('compose.yml') ||
  !ui.includes('client_max_body_size 0;') ||
  !ui.includes('proxy_set_header Upgrade $http_upgrade;') ||
  !ui.includes('BEGIN PIPELINE-GENERATOR MANAGED ROUTES') ||
  !ui.includes('multiple HTTPS server blocks') ||
  !ui.includes('`/${nginxDirectory}/${projectHost}-${normalizedEnvironment}.conf`') ||
  !ui.includes('Nginx and Compose files below')
) {
  fail('Step 1 must provision the Docker/Nginx repositories and idempotent starter configuration files.');
}
if (!ui.includes("return `${normalizePart(service, 'Service name')} ${normalizePart(environment, 'Environment')}`;")) {
  fail('Classic Release names must contain only uppercased service and environment.');
}
if (ui.includes('state.repositoryName = repo.name')) {
  fail('generated repository metadata must not overwrite the source repository identity used for retries.');
}
if (!ui.includes('state.generatedRepositoryName = repo.name')) {
  fail('generated repository metadata must be stored separately from source repository state.');
}
if (ui.includes('const updateUrl = buildPipelinesApiUrl')) {
  fail('Azure DevOps Server does not support updating Pipeline definitions through PUT /_apis/pipelines/{id}.');
}
if (!ui.includes("method: 'PUT'") || !ui.includes('yamlFilename: desiredConfig.path')) {
  fail('Pipeline reconciliation must use GET-modify-PUT through the Build Definitions API.');
}
if (!ui.includes('normalizeComparableFolder') || !ui.includes('.toLowerCase()')) {
  fail('Pipeline and Release folder comparisons must tolerate server-normalized casing.');
}
if (!ui.includes("conditions: [{ name: 'ReleaseStarted', conditionType: 'event', value: '', result: null }]")) {
  fail('classic Release environments must start when the release is created.');
}
if (!ui.includes("approvals: [{ rank: 1, isAutomated: true, isNotificationOn: false }]")) {
  fail('classic Release environments must include automated pre/post approvals for Azure DevOps Server.');
}
if (!ui.includes("executionOrder: 'beforeGates'") || !ui.includes("executionOrder: 'afterSuccessfulGates'")) {
  fail('classic Release pre/post approval execution order is incomplete.');
}
if (
  !ui.includes("'$expand': 'Artifacts'") ||
  !ui.includes("artifactSourceId: `${projectId}:${pipelineId}`") ||
  !ui.includes('releaseDefinitionMatches')
) {
  fail('classic Release definitions must be found/reconciled by Pipeline artifact as well as exact name.');
}
if (!ui.includes('variableGroups: [Number(variableGroup.id)]') || !ui.includes('mergedVariableGroups')) {
  fail('classic Release definitions must link KomodoAPI at definition scope and preserve existing groups.');
}

const restContracts = read('docs/rest-api-contracts.md');
if (!restContracts.includes("`7.1-preview.1`") || !restContracts.includes('`repositoryId` query parameter')) {
  fail('REST documentation must describe the browser Pipeline API version and repositoryId binding.');
}
if (!restContracts.includes('default is `7.1`')) {
  fail('REST documentation must distinguish the shell API_VERSION default from the browser contract.');
}
if (!restContracts.includes('GET-modify-PUT') || !restContracts.includes('exact generated YAML filename')) {
  fail('REST documentation must describe Build Definition reconciliation and filename-based Pipeline naming.');
}
if (
  !restContracts.includes('`_signout`') ||
  !restContracts.includes('`_settings/extensions?tab=installed`') ||
  !/does not\s+accept a PAT/.test(restContracts)
) {
  fail('REST documentation must describe the full-session reauthentication path and PAT prohibition.');
}

console.log('Extension validation passed: manifest scopes, release config, and provisioning assets are consistent.');
