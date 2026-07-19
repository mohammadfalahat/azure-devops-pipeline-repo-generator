/*
 * Classic Release definition settings for Pipeline Generator.
 *
 * This file is packaged with the extension and is deliberately free of tokens,
 * passwords, and other credentials. All REST calls use the short-lived Azure
 * DevOps access token issued to the signed-in user by VSS.getAccessToken().
 *
 * To use your existing deployment script, either replace `content` below, or
 * change `scriptSource` to `azureReposFile`. The source repository must be in
 * the same Azure DevOps collection and the person running the generator must
 * have permission to read it.
 */
window.PipelineGeneratorReleaseConfig = Object.freeze({
  enabled: true,
  folder: '\\komodo',
  environmentName: 'komodo',
  nameSuffix: '_Release',
  bashTaskName: 'Run Komodo deployment',
  scriptSource: {
    type: 'inline',
    content: `#!/usr/bin/env bash
set -euo pipefail

# Replace this placeholder with your predefined Komodo deployment commands.
echo "Komodo release job started."
echo "Build artifact directory: \${SYSTEM_ARTIFACTSDIRECTORY:-not-set}"
`

    // To load the Bash content from an Azure Repos file instead, replace the
    // block above with this (keep it in the same collection):
    // type: 'azureReposFile',
    // project: 'Tools',
    // repository: 'deployment-scripts',
    // branch: 'main',
    // path: '/komodo/release-task.sh'
  }
});