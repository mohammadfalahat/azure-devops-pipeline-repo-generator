# Development and operations

This guide covers local development, configuration, testing, packaging,
installation, shell automation, live verification, and troubleshooting.

## Prerequisites

For extension development:

- Node.js 18 or newer;
- npm;
- network access to the target Azure DevOps host for live tests;
- `tfx-cli`, invoked through `npx` by the package scripts.

For `scripts/provision-pipeline-release.sh`:

- Bash;
- `curl`;
- `jq`;
- Git only when `RELEASE_BASH_SCRIPT_GIT_URL` mode is used.

The target Azure DevOps project must already contain or authorize:

- `SharedTemplates/SharedTemplates`;
- `/pipeline-generator.yml` on that repository's `main` branch, with non-empty
  Environment `name`/`domain` records; a legacy `servers` list may remain but
  is ignored;
- `/komodo-servers-creds.env` on the same repository/branch, containing the
  centrally managed Komodo address and Server-Read API credential;
- service endpoint `ShonizCollection` for that repository resource;
- template `build-push-komodo.yml` on its `main` branch;
- variable group `KomodoAPI` with secret variables `AZP_TOKEN`,
  `KOMODO_API_KEY`, and `KOMODO_API_SECRET`; the extension identity/user must
  be able to resolve and use the group;
- the selected Docker Registry service connection;
- the selected agent queue.

These external resources are referenced but never created by this extension.

## Local setup

```bash
npm ci
npm test
```

The project is a static extension. `npm run build` is a documented no-op:

```bash
npm run build
```

There is no `src`-to-`dist` compilation. Edit runtime files in `dist/` directly
and review them as shipped code.

The manifest declares a Cloud target, but the full workflow is live-verified
only on Azure DevOps Server. Before claiming Azure DevOps Services support,
test and, if required, implement the separate `vsrm.dev.azure.com` Release API
base URI.

## Configuration surfaces

### Extension manifest

Edit `vss-extension.json` when changing:

- extension ID, publisher, display name, description, or version;
- contribution target menus or registered object ID;
- addressable/package files;
- Azure DevOps Services/Server target range;
- requested token scopes;
- gallery metadata.

The manifest contribution ID and `registeredObjectId` must remain synchronized
with `VSS.register('generate-pipeline-action', ...)` in `dist/menu-action.js`.
Any new runtime asset referenced by HTML or JavaScript must be included under a
packaged path.

### Form and provisioning defaults

`dist/ui.js` currently owns:

- `defaultValues` and default Pool/registry options;
- the `SharedTemplates/SharedTemplates:/pipeline-generator.yml@main` deployment
  environment/domain source and its small strict parser;
- strict parsing of the central Komodo credential file and direct
  `ListFullServers` response;
- environment-to-server label inference;
- scaffold branch `main`;
- Pipeline folder `\komodo`;
- Git, Pipeline, and Release API versions;
- Bash task ID;
- YAML naming and rendering;
- REST reconciliation logic.

`dist/index.html` contains disabled loading placeholders for Environment and
Komodo Server. Update `pipeline-generator.yml` to change environments. Server
enable/disable changes come directly from Komodo and require no repackaging.
A missing/unreadable/invalid environment or credential file, rejected Komodo
request, blocked CORS preflight, or empty enabled-server result is blocking and
has no compiled-in fallback.

Use structured Environment records:

```yaml
environments:
  - name: pro
    domain: bulutcom.cloud
  - name: dev
    domain: bulutdev.ir
```

The compact `"dev:bulutdev.ir"` form is parsed for migration, but structured
records are the maintained contract. Every record needs a domain because the
Nginx starter host and certificate filenames are derived from it.

### Central Komodo credential file

Create `/komodo-servers-creds.env` on `main` in
`SharedTemplates/SharedTemplates` with exactly these fields:

```dotenv
KOMODO_ADDRESS=https://komodo.buluttakin.com
KOMODO_API_KEY=<server-read-key>
KOMODO_API_SECRET=<server-read-secret>
```

The same single file serves every project; colleagues do not define per-project
discovery credentials. The API user must be enabled, have base `Read` only for
the Server resource type, and have `None` for other resource types and create
permissions. By operator decision this dedicated credential is browser-readable
and not treated as a write-capable secret. Do not reuse a deployment/admin key.

The browser reads this file using its current Azure DevOps extension token,
calls Komodo `/read` with `ListFullServers`, immediately projects the response
to enabled non-template server names, and does not persist or log credential
values. Komodo must allow `https://azure.buluttakin.com` in
`KOMODO_CORS_ALLOWED_ORIGINS` when its allowed-origin list is non-empty. The
direct call uses custom `X-Api-Key` and `X-Api-Secret` headers, so verify both
the OPTIONS preflight and POST response in the browser.

### Release definition configuration

`dist/release-config.js` is loaded before `ui.js` and assigns
`window.PipelineGeneratorReleaseConfig`.

| Field | Type | Behavior |
| --- | --- | --- |
| `enabled` | Boolean | `false` skips Release creation but still creates YAML/Pipeline |
| `folder` | String | Normalized to an Azure DevOps folder beginning with `\` |
| `environmentName` | String | Required name of the single classic Release environment |
| `bashTaskName` | String | Display name of the Bash@3 task; default `Run Komodo deployment` |
| `variableGroupName` | String | Exact project Variable Group name; default and supported deployment contract is `KomodoAPI` |
| `requiredVariableNames` | String array | Names that must exist before a Release write: `AZP_TOKEN`, `KOMODO_API_KEY`, `KOMODO_API_SECRET` |
| `scriptSource` | Object | Packaged asset, non-empty literal content, or an Azure Repos file reference |

Packaged-file mode is the default:

```js
scriptSource: {
  type: 'packagedFile',
  path: 'release-inline-task.sh'
}
```

The installed asset is fetched when the definition is generated or reconciled,
then stored as the Bash task's inline `inputs.script` text. Edit
`dist/release-inline-task.sh` and
`scripts/release-inline-task.example.sh` together; the validator requires them
to be byte-equivalent after trailing whitespace normalization.

Inline mode:

```js
scriptSource: {
  type: 'inline',
  content: `#!/usr/bin/env bash
set -euo pipefail

echo "Run deployment here"
`
}
```

Azure Repos file mode:

```js
scriptSource: {
  type: 'azureReposFile',
  project: 'Tools',
  repository: 'deployment-scripts',
  branch: 'main',
  path: '/komodo/release-task.sh'
}
```

File mode reads the file at definition-generation time and embeds its text into
the resulting inline Release task. Later repository changes do not update the
Release until the generator runs again; the next run compares and updates the
same-name or same-Pipeline Release definition.

Never place credential values in this config or packaged script. Azure DevOps
macro names such as `$(AZP_TOKEN)` are expected placeholders, not credentials.
The validator checks the packaged/documented script pair and ensures the
secret-derived Authorization header is neither traced nor placed in process
arguments, but it is not a general-purpose secret scanner.

### Generated YAML template

Change `buildPipelineYaml` in `dist/ui.js` when modifying repository resources,
variable groups, shared template path, or parameters. Because the YAML is
assembled as text:

- preserve indentation explicitly;
- add escaping or validation before accepting new free-text values;
- run `npm test` and inspect a concrete generated document;
- validate the document in a test Pipeline on the target server.

`buildPipelineFilename` requires Environment and returns
`<project>-<repository>-<SanitizedBranch>To<UPPERCASE-ENVIRONMENT>.yml`;
`buildPipelineName` returns that filename unchanged. `buildReleaseName` uses
only uppercased Service and Environment values. If either naming contract
changes, update migration lookups, artifact aliases, tests, and documentation
together.

### Bundled SDK

The SDK files under `dist/lib/` are intentionally packaged for on-premises
compatibility. If updating `vss-web-extension-sdk`:

1. update the npm dependency;
2. replace both bundled SDK artifacts from the installed package;
3. test the action contribution and both `pipeline-generator-dialog` and
   `pipeline-generator-hub` host iframes on the oldest supported server;
4. verify no platform SDK URL introduces an authentication prompt;
5. package and inspect the VSIX contents.

## Offline test suite

Run:

```bash
npm test
```

This executes five independent checks.

### Extension validation

`scripts/validate-extension.js`:

- parses the manifest and checks required scopes;
- evaluates `release-config.js` in a VM;
- checks Release config, `KomodoAPI` requirements, packaged Bash parity, and
  Authorization-header xtrace safety;
- verifies load order of Release config before UI;
- checks the Dialog control and Azure Repos Hub manifest contracts;
- checks for required Pipeline URL binding and Release approval contracts.

Most UI checks are source-string assertions. The validator does not execute the
DOM workflow, mock REST responses, or parse generated YAML.

### Action behavior regression test

`scripts/validate-action-behavior.js` instruments `dist/menu-action.js` only in
a Node VM. It verifies the fully-qualified Dialog and Hub contribution IDs,
host service IDs, token-free branch/repository configuration, successful
Dialog launch, and in-host Hub navigation when Dialog is unavailable. It also
asserts the action never requests a token or opens a detached window. No
runtime test hook is written to `dist/`.

### UI behavior regression test

`scripts/validate-ui-behavior.js` instruments the production `dist/ui.js` IIFE
in memory and exposes selected functions only inside a Node VM. It uses mocked
REST responses to assert that:

- host extension tokens produce Bearer headers and the browser runtime has no
  PAT credential path;
- session restart clears the in-memory host token and asks the Azure DevOps host
  navigation service to open the collection-relative `_signout` route;
- `pipeline-generator.yml` is parsed into dynamic Environment/domain options
  and read from the exact shared repository path;
- the exact central credential-file Git URL is called with the ADO Bearer token;
- the subsequent direct Komodo request uses the parsed synthetic credentials
  in tests, filters disabled/template records, and keeps only server names;
- Docker/Nginx DevOps repository names, Compose/shared-Nginx paths, project
  hostname, service route, ports, WebSocket headers, and certificate paths
  follow the documented convention; managed routes always use Docker DNS
  `resolver 127.0.0.11 ipv6=off` and `set $target <container>`; root uses a
  trailing-slash proxy target, while non-root uses `/<service>/`, no rewrite,
  and a no-URI-slash proxy target so the request URI is preserved; root is
  ordered last and older managed paths/targets/generated rewrites are migrated
  while only a missing Location is inserted;
- Pipeline name exactly equals the Branch-to-Environment transition YAML filename;
- Release name contains only uppercased Service and Environment;
- an existing byte-identical YAML file is read and reused without a Git Push or
  no-op commit;
- a sparse exact-name Pipeline reference is resolved through the complete Build
  Definition, and a correctly linked definition with folder `\KOMODO` is reused
  without any PUT or revision increment;
- a 0.1.37 Environment-first or earlier branch-only Pipeline is found by legacy
  name or YAML path, selected deterministically, and migrated to the
  BranchToEnvironment name/path through Build Definitions GET-modify-PUT while
  preserving its ID;
- no `PUT /_apis/pipelines/{id}` is sent;
- a legacy Release is found by `<projectId>:<pipelineId>` Build artifact source,
  updated in place, and retains its revision, environment ID, and deployment
  phase ID;
- `KomodoAPI` is queried with `actionFilter=Use`, its required variable names
  are validated, and its ID is added at Release-definition scope while existing
  additional group IDs are preserved.

The production file is not modified and contains no test-only global hook.

### Service Hook self-test

`scripts/service-hook-listener.js --self-test` checks that the payload summarizer
extracts event, project, repository, and collection values from a sample push.

### Shell payload self-test

`scripts/provision-pipeline-release.sh --self-test` uses `jq` to build sample
Pipeline and Release payloads, then asserts folder, YAML path, artifact, Bash
task, condition, automated approval fields, and the definition-level Variable
Group ID. It performs no network calls.

### What offline tests do not prove

The suite does not prove:

- extension installation or host authorization;
- branch-menu contribution visibility;
- VSS session-token issuance;
- full browser-session sign-out and subsequent interactive login;
- Azure DevOps Server API compatibility;
- YAML existence or validity in Azure Repos;
- actual Pipeline/Release definition linkage;
- Pipeline execution, template authorization, service connections, variable
  groups, agent availability, or deployment success.

Use the live verification checklist below for those concerns.

## Packaging

Run tests and then package:

```bash
npm test
npm run extension:package
```

The npm command executes:

```text
tfx extension create --manifest-globs vss-extension.json --rev-version
```

`--rev-version` can modify the manifest version. After packaging:

1. inspect `git diff -- vss-extension.json`;
2. confirm the intended version;
3. list or unzip the VSIX and verify the manifest, `dist/`, SDK, icon, and
   README are present;
4. compute and record SHA-256 for the exact artifact being tested;
5. rerun `npm test` if the version or configuration was adjusted.

Do not treat an older same-name VSIX as the new artifact. Record the exact file
and hash in the E2E state when it is installed.

## Publishing and installation

The preferred on-premises path is manual upload to the server gallery, followed
by installation/enabling for the target collection:

1. Open the Azure DevOps extension management/gallery page.
2. Upload the new VSIX.
3. Install or update it for the intended collection.
4. Confirm the installed version and that install-state flags are clear.
5. hard-refresh Azure Repos and verify **Generate pipeline** appears on a branch
   menu.

`npm run extension:publish -- <vsix>` invokes `tfx-cli` and expects
`ADO_SERVICE_URL` and `AZP_TOKEN`. Use a separate, short-lived
extension-management PAT. Never store it in shell history, source files,
`.env`, npm configuration, documentation, or the VSIX. Be aware that the
current tfx script expands `--token` for the child process; prefer manual upload
when local process-argument visibility is unacceptable.

A PAT that can call project or extension-management REST APIs may still receive
401 from Marketplace-oriented `tfx extension show/publish`. Treat gallery
publishing rights and collection installation rights as separate capabilities.

### Stale installation recovery

When install/update reports a conflict but the normal by-name endpoint reports
not found:

1. list installed extensions with disabled and error entries included;
2. inspect `installState.flags` and installed version;
3. if the extension is stuck in `error` or `needsReauthorization`, remove that
   collection installation;
4. clean-install the intended version with an empty POST body;
5. verify `installState.flags` becomes `none` before browser testing.

Use the exact verified sequence in `azure-devops-e2e-context.md`; do not delete
healthy installations as a first diagnostic step.

## Terminal provisioning fallback

The shell script is not equivalent to the full browser workflow. It does not
create a repository, generate YAML, push a file, or set a default branch. It
starts only after the target YAML already exists.

It performs:

1. local command/environment validation;
2. Bash script-source loading;
3. queue ID resolution when only a queue name is supplied;
4. project ID lookup;
5. `KomodoAPI` ID resolution and required-variable-name validation;
6. target repository ID lookup;
7. Pipeline lookup by exact name and creation if absent;
8. Release definition lookup by exact name and creation if absent;
9. optional Release instance creation.

Queue resolution and Bash source loading occur before checking whether a
same-name Release definition already exists. Therefore even a reuse-only run
still requires a valid queue input and readable, non-empty Bash source.

### Required shell variables

| Variable | Meaning |
| --- | --- |
| `AZP_TOKEN` | PAT used through curl Basic authentication |
| `ADO_URL` | Server root, without collection |
| `COLLECTION` | Collection name/path segment |
| `PROJECT` | Project name/path segment |
| `PIPELINE_NAME` | Exact Pipeline lookup/create name |
| `REPO_NAME` | Existing target repository that contains the YAML, not the source app repository |
| `YAML_PATH` | Exact existing YAML path, preferably beginning with `/` |

For parity with the browser's naming contract, set `PIPELINE_NAME` to the
basename of `YAML_PATH`, including `.yml`. The shell accepts other names for
backward compatibility and does not enforce this equality.

Supply either `RELEASE_AGENT_QUEUE_ID` or `RELEASE_AGENT_QUEUE_NAME`. Also
supply one Bash source:

- `RELEASE_BASH_SCRIPT`;
- `RELEASE_BASH_SCRIPT_FILE`;
- `RELEASE_BASH_SCRIPT_GIT_URL` plus `RELEASE_BASH_SCRIPT_GIT_PATH`.

Source priority is inline environment value, local file, then Git clone. Git
mode defaults `RELEASE_BASH_SCRIPT_GIT_REF` to `main`. Use a credential helper
for private Git access; never embed a PAT in the clone URL.

### Optional shell variables

| Variable | Default |
| --- | --- |
| `DEFAULT_BRANCH` | `refs/heads/main` |
| `PIPELINE_FOLDER` | `komodo`, normalized to `\komodo` |
| `RELEASE_NAME` | `${PIPELINE_NAME}_Release` |
| `RELEASE_FOLDER` | `komodo`, normalized to `\komodo` |
| `RELEASE_ENVIRONMENT_NAME` | `komodo` |
| `RELEASE_VARIABLE_GROUP_NAME` | `KomodoAPI` |
| `RELEASE_ARTIFACT_ALIAS` | `_${PIPELINE_NAME}` |
| `RELEASE_BASH_TASK_NAME` | `Run inline Bash` |
| `RELEASE_BASH_SCRIPT_GIT_REF` | `main` |
| `CREATE_RELEASE_INSTANCE` | `false` |
| `API_VERSION` | `7.1` |
| `RELEASE_API_VERSION` | `7.1-preview.4` |

### Safe invocation pattern

Create the token file interactively outside the repository:

```bash
umask 077
read -rsp 'Azure DevOps PAT: ' CODEX_UI_PAT
printf '%s' "$CODEX_UI_PAT" > "/tmp/codex-azp-token-$(id -u)"
unset CODEX_UI_PAT
echo
```

Then run the script without placing the PAT literal in history:

```bash
AZP_TOKEN="$(<"/tmp/codex-azp-token-$(id -u)")" \
ADO_URL='https://azure.example.local' \
COLLECTION='DefaultCollection' \
PROJECT='ExampleProject' \
PIPELINE_NAME='ExampleProject_Api_demo' \
REPO_NAME='ExampleProject_Azure_DevOps' \
YAML_PATH='/exampleproject-api-feature-demo.yml' \
DEFAULT_BRANCH='refs/heads/main' \
PIPELINE_FOLDER='komodo' \
RELEASE_AGENT_QUEUE_NAME='PublishDockerAgent' \
RELEASE_BASH_SCRIPT_FILE='scripts/release-inline-task.example.sh' \
CREATE_RELEASE_INSTANCE='false' \
bash scripts/provision-pipeline-release.sh
```

Remove the token file when finished:

```bash
shred -u "/tmp/codex-azp-token-$(id -u)"
```

On an internal host, configure `NO_PROXY` or unset proxy variables for the
process. Trust the internal CA rather than editing the script to disable TLS
verification.

### Shell reconciliation differences

The shell script is idempotent only by exact name:

- an existing Pipeline is reused without checking/updating its YAML path,
  repository, branch, or folder;
- an existing Release definition is reused without checking/updating its
  artifact, queue, script, Variable Groups, approvals, environment, or folder.

The browser path performs Pipeline and Release reconciliation; the shell path
does not. Operators must read back existing objects before assuming the
requested arguments are now applied.

The shell URL builder inserts `COLLECTION` and `PROJECT` directly into paths and
does not URL-encode them. Names containing spaces or reserved URL characters
require a code fix rather than ad-hoc shell escaping. Like the browser path, the
shell does not paginate list results.

Set `CREATE_RELEASE_INSTANCE=true` only when an actual Release run is intended.
It creates a Release using the latest artifact and is a state-changing execution
operation, not merely definition provisioning.

## Development Service Hook listener

Start the listener with:

```bash
npm run service-hook:listen
```

Options:

```bash
npm run service-hook:listen -- --port 8081
npm run service-hook:listen -- --quiet
LOG_PAYLOADS=false npm run service-hook:listen
```

The listener is unrelated to extension runtime provisioning. It accepts any
HTTP request, buffers the entire body, attempts JSON parsing, logs metadata and
raw payload by default, and always returns HTTP 200 JSON. It has no
authentication, signature validation, route restriction, body-size limit,
durable storage, or retry handling.

Use it only as a local development aid. Do not expose it publicly or use it in
production without adding transport security, authentication, request limits,
event validation, redaction, and controlled logging.

## Live E2E verification checklist

Before testing:

1. read `azure-devops-e2e-state.yaml` and preserve recorded live resources;
2. run `npm test`;
3. confirm the VSIX hash and installed version;
4. confirm extension install-state flags and project enablement;
5. choose a disposable branch/project or explicitly agree on reuse behavior;
6. take API snapshots of existing repository, YAML, Pipeline, and Release
   definitions;
7. state whether a run or Release instance is forbidden or requested.

During the browser test:

1. open the selected branch's menu and choose **Generate pipeline**;
2. confirm the generator is either a modal or the Pipeline Generator Hub under
   Azure Repos, never a separate browser tab;
3. verify project, source repository, source branch, target repository, and
   inferred form values, including the dynamically loaded deployment targets;
4. confirm the generator obtained a Bearer host token through its in-frame SDK
   handshake without asking for a PAT;
5. submit once and record the five status transitions;
6. confirm success stays on the form and displays working Nginx, Compose, and
   Pipeline links without queueing a run;
7. do not infer success solely from the UI message.

After the test, read back and compare:

- exact YAML path and branch;
- Docker/Nginx DevOps repositories, `/environments`, and selected Environment
  Compose/Nginx starter files;
- Pipeline ID, name, folder, repository ID, default branch, and YAML path;
- classic Build Definition `process.yamlFilename`;
- Release ID/name/folder, artifact Pipeline and repository IDs;
- environment, queue ID/name, Bash task/version/target, event condition, and
  automated approvals;
- whether a run or Release instance was created.

Record the authentication method precisely. A PAT-adapted headless test proves
different boundaries than a normal signed-in session that successfully calls
`VSS.getAccessToken()`.

Update both E2E context files after any live change and identify any disposable
resources that were cleaned up.

## Troubleshooting guide

| Symptom | Likely cause | First checks |
| --- | --- | --- |
| Action missing from branch menu | Extension not enabled, stale browser assets, or menu target mismatch | Confirm installed version/state, hard-refresh, inspect manifest contribution targets |
| Session-token HTTP 401/403 or an expired-session error | The current browser sign-in is stale or the user lost project access | Use **Sign out and authenticate again**, complete the full login flow, then reopen the branch action |
| Generator opens in a separate tab | Version 0.1.25 or older is still served, or stale action assets remain cached | Verify installed/asset version 0.1.41 and hard-refresh; current code has no detached-window path |
| Generator opens as an Azure Repos Hub instead of a modal | This Azure DevOps Server does not expose the custom Dialog service | Expected compatibility behavior; the Hub is still a host iframe |
| Hub width repeatedly shrinks while blank space grows on the right | Legacy `VSS.resize()` used the form's changing `scrollWidth`, creating a host/iframe width feedback loop | Install 0.1.41; it never calls host resize and keeps contribution width fixed |
| Hub form is clipped or mouse-wheel scrolling does nothing | Legacy iframe root/body scrolling is suppressed by the host | Version 0.1.41 uses a fixed full-viewport `.wrapper` as an explicit scroll container; verify the served asset version and hard-refresh |
| `HostAuthorizationNotFound` inside Dialog/Hub | Collection installation has no authorization record for the extension scopes, or that record is stale | Select **Open extension authorization**; a Collection Administrator must authorize Pipeline Generator in Collection Settings → Extensions. If no action exists, reinstall the same published version |
| Generator says another access-token error | Page opened directly, hosted iframe SDK handshake failed, or host denied token | Launch from a branch action; retry full sign-in only after extension authorization is confirmed |
| Browser displays Basic login prompts | Platform SDK/API request received an auth challenge | Confirm bundled SDK is used, Bearer token is present, and fed-auth redirects are suppressed |
| Repository created but YAML push fails | Contribute/default branch permission or stale old object ID | Inspect Git API response and current `main` ref |
| `YamlFileNotFoundException` | Path/branch/repository mismatch at Pipeline create time | Read exact YAML from generated repo `main`; ignore Pipeline folder when checking path |
| Pipeline request returns 401/403/TF400813 | Missing Pipeline Create/Edit permission or folder ACL | Check Project settings → Pipelines → Security and `\komodo` folder security |
| Pipeline create fails despite repository in JSON body | Missing on-prem `repositoryId` query parameter or wrong API contract | Inspect final URL and generated repository GUID |
| HTTP 2xx but no Pipeline ID | Unexpected server/proxy response shape | Inspect response safely; the UI intentionally rejects it |
| Release fails with `VS402877` | Empty/missing pre/post approvals | Keep automated approvals and correct execution orders in both payload builders |
| Environment stays unavailable | `pipeline-generator.yml` is missing/invalid, an Environment lacks a valid domain, `main` is absent, or the user cannot read `SharedTemplates/SharedTemplates` | Verify the exact file URL, structured non-empty `environments` records, and repository Read permission; there is no static fallback |
| Komodo Server remains on Loading/unavailable | Central credential file is missing/invalid/unreadable, Komodo CORS blocks the ADO origin/custom headers, Komodo rejects the read credential, or no visible Server has `config.enabled: true` | Verify the exact SharedTemplates file path/branch and Read permission, inspect OPTIONS/POST status without logging header values, confirm `KOMODO_CORS_ALLOWED_ORIGINS`, and test `ListFullServers` with the dedicated user |
| Step 1 fails after creating the Azure DevOps repository | Docker/Nginx support repository creation, bootstrap push, or shared Nginx merge was denied/ambiguous | Grant Create repository/Contribute permission; for Nginx also verify balanced braces, complete managed markers, and exactly one matching port-443 server block before rerunning |
| Release Step 5 gets 401 while resolving a queue | Extension token lacks `vso.agentpools`, has not been reauthorized after adding it, or the user cannot view/use the queue | Verify installed version 0.1.41 scopes, reauthorize/reinstall it, then confirm the current user can read and Use the queue |
| Registry choices fall back to defaults with a 401 in Console | Extension token lacks `vso.serviceendpoint` or has not been reauthorized | Authorize the updated 0.1.41 scopes; the extension only reads endpoint names/types |
| Release Step 5 cannot resolve `KomodoAPI` | The extension lacks `vso.variablegroups_read`, the new scope has not been authorized, or the group/user lacks Use permission | Reauthorize/install 0.1.41, confirm `KomodoAPI` exists in the current project, and grant the user/extension permission to use it |
| Release reports missing required variables | `KomodoAPI` exists but lacks one of the wrapper inputs | Add secret variables `AZP_TOKEN`, `KOMODO_API_KEY`, and `KOMODO_API_SECRET` with exact casing; do not put their values in source control |
| Release cannot be created after a Pipeline error | Step 5 never ran because Step 4 aborted | Fix/read back Pipeline first; then rerun so Release migration/create can execute |
| Release has a legacy name/configuration | It was created by an older extension | Version 0.1.27 finds it by Pipeline artifact ID and reconciles it without deleting it |
| Pipeline update returns HTTP 405 for PUT | The server does not support PUT on `/_apis/pipelines/{id}` | Use the implemented Build Definition GET-modify-PUT path, not Pipelines PUT |
| Pipeline is rewritten because folder casing differs | Server normalized `\komodo` to another casing | Version 0.1.27 compares folder casing case-insensitively |
| An unchanged Pipeline gains a revision on every rerun | Pipelines by-ID omitted `repository.defaultBranch`, so the sparse response looked mismatched | Version 0.1.31 compares the complete Build Definition and skips PUT when the binding already matches |
| Object exists but name lookup misses it | It is beyond the first API page or a duplicate exists in another folder | Add continuation-token pagination and explicit disambiguation |
| Dockerfile is not auto-detected | No Dockerfile exists on the selected source branch or the tree read failed | Enter the path manually and inspect the Git Items response |
| Terminal calls return nginx 403 | Internal Azure DevOps request went through environment proxy | Configure `NO_PROXY`/unset proxy variables for the host |
| `DELETE /_apis/pipelines/{id}` returns 405 | Endpoint unsupported for deletion on this server | Delete disposable Pipeline via Build Definitions API |
| Page redirects after success, leaves the completed form active, or does not show three review links | Stale extension version/assets | Verify installed asset version 0.1.41 and hard-refresh; current flow locks/collapses the form and focuses the Nginx, Compose, and Pipeline links |

For environment-specific IDs, exact recovery endpoints, failed automation
approaches, and the last verified successful resource graph, use
`azure-devops-e2e-context.md` rather than repeating exploratory tests.

## Change checklist

Before merging or publishing a change:

1. identify whether the change affects action context, token transfer, YAML,
   Git writes, Pipeline binding, Release payload, naming, or package manifest;
2. update browser and shell implementations when they share a contract;
3. update or add offline assertions;
4. update architecture, REST, and operations documentation;
5. run `npm test` and `git diff --check`;
6. inspect for secrets and internal temporary evidence;
7. package and record artifact version/hash when release testing is requested;
8. perform proportional live readback verification;
9. record test method, resulting IDs, and cleanup status in the E2E context.
