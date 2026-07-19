#!/usr/bin/env node

/*
 * Offline validation for the static extension assets. It deliberately does not
 * contact Azure DevOps; REST/API behavior must be verified against the target
 * Azure DevOps Server after publishing the VSIX.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const read = (relativePath) => fs.readFileSync(path.join(root, relativePath), 'utf8');
const fail = (message) => {
  throw new Error(`[extension validation] ${message}`);
};

const manifest = JSON.parse(read('vss-extension.json'));
const requiredScopes = ['vso.code', 'vso.code_manage', 'vso.build', 'vso.build_execute', 'vso.release', 'vso.release_manage'];
for (const scope of requiredScopes) {
  if (!manifest.scopes?.includes(scope)) {
    fail(`manifest is missing required scope: ${scope}`);
  }
}

const releaseConfigContext = { window: {} };
vm.runInNewContext(read('dist/release-config.js'), releaseConfigContext, { filename: 'dist/release-config.js' });
const releaseConfig = releaseConfigContext.window.PipelineGeneratorReleaseConfig;
if (!releaseConfig?.enabled || releaseConfig.folder !== '\\komodo') {
  fail('release configuration must enable the \\komodo release folder.');
}
const scriptSource = releaseConfig.scriptSource || {};
if (scriptSource.type === 'inline') {
  if (!scriptSource.content?.trim()) {
    fail('release configuration must provide a non-empty inline Bash script.');
  }
  if (/AZP_TOKEN|PERSONAL_ACCESS_TOKEN|\bPAT\b/i.test(scriptSource.content)) {
    fail('release inline script must not embed an Azure DevOps token.');
  }
} else if (scriptSource.type === 'azureReposFile') {
  for (const field of ['project', 'repository', 'path']) {
    if (!String(scriptSource[field] || '').trim()) {
      fail(`azureReposFile script source requires ${field}.`);
    }
  }
} else {
  fail('release scriptSource.type must be inline or azureReposFile.');
}

const html = read('dist/index.html');
if (!/release-config\.js[\s\S]*ui\.js/.test(html)) {
  fail('index.html must load release-config.js before ui.js.');
}

const ui = read('dist/ui.js');
for (const requiredImplementation of ['ensureReleaseDefinition', 'buildReleaseDefinitionPayload', 'BASH_TASK_ID', 'runProvisioningStep']) {
  if (!ui.includes(requiredImplementation)) {
    fail(`ui.js is missing ${requiredImplementation}.`);
  }
}

console.log('Extension validation passed: manifest scopes, release config, and provisioning assets are consistent.');