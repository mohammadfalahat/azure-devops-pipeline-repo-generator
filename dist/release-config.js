/*
 * Classic Release definition settings for Pipeline Generator.
 *
 * This file is packaged with the extension and is deliberately free of tokens,
 * passwords, and other credentials. All REST calls use the short-lived Azure
 * DevOps host token issued to the signed-in user by VSS.getAccessToken().
 *
 * The packaged Bash wrapper is loaded at definition-generation time and stored
 * as inline text in the classic Release task. No credential values are stored
 * in this file or in the packaged script; Azure DevOps expands variables from
 * the linked KomodoAPI variable group at Release execution time.
 */
window.PipelineGeneratorReleaseConfig = Object.freeze({
  enabled: true,
  folder: '\\komodo',
  environmentName: 'komodo',
  bashTaskName: 'Run Komodo deployment',
  variableGroupName: 'KomodoAPI',
  requiredVariableNames: Object.freeze(['AZP_TOKEN', 'KOMODO_API_KEY', 'KOMODO_API_SECRET']),
  scriptSource: {
    type: 'packagedFile',
    path: 'release-inline-task.sh'

    // To load different Bash content from Azure Repos instead, replace the
    // block above with this (keep it in the same collection):
    // type: 'azureReposFile',
    // project: 'Tools',
    // repository: 'deployment-scripts',
    // branch: 'main',
    // path: '/komodo/release-task.sh'
  }
});
