# Architecture and runtime flow

This document describes version 0.1.41 from the implementation in
`vss-extension.json`, `dist/menu-action.js`, `dist/ui.js`, and
`dist/release-config.js`.

## System boundary

Pipeline Generator is a client-only Azure DevOps web extension. It has:

- no application server;
- no database;
- no persisted runtime credential;
- no build-time generated JavaScript;
- no background worker or webhook dependency.

All Azure DevOps provisioning calls are made directly from the user's browser
to the same collection that hosts the extension. The security principal is the
signed-in user represented by the token returned by `VSS.getAccessToken()`.
The manifest includes separate read scopes for agent pools/queues, service
endpoints, and Variable Groups; Build and Release scopes do not implicitly
grant those resource-area permissions. The browser also reads the central
`/komodo-servers-creds.env` file from `SharedTemplates`, then calls the Komodo
1.19.x read API using the operator-designated non-confidential Server-Read
credential. Credential values remain in page/request memory only and are never
logged or persisted. Komodo must allow the Azure DevOps origin through CORS.

The separate shell script is an operational fallback, not a backend for the
extension. It receives a PAT and assumes the target repository and YAML file
already exist.

## Component model

| Component | Loaded by | Main responsibilities |
| --- | --- | --- |
| `vss-extension.json` | Azure DevOps extension host | Declares the branch-menu action, Dialog control, Azure Repos Hub, supported hosts, addressable files, and token scopes |
| `menu-action.html` / `menu-action.js` | Hidden action contribution iframe | Initializes VSS SDK, registers `generate-pipeline-action`, extracts branch context, warms assets, and asks the host to open the form |
| `index.html` / `ui.js` | Dialog or `pipeline-generator-hub` host iframe | Reads host configuration/navigation state, obtains the current-user host token, hydrates the form, performs five provisioning steps, displays errors, and renders Nginx/Compose/Pipeline review links |
| `release-config.js` | Loaded before `ui.js` | Exposes immutable `window.PipelineGeneratorReleaseConfig` with Release settings, `KomodoAPI` requirements, and Bash source selection |
| `release-inline-task.sh` | Fetched by `ui.js` from the installed extension assets | Provides the wrapper text embedded into the classic Release Bash task |
| `VSS.SDK*.js` | Action and generator pages | Supplies the legacy VSS extension APIs required by the supported on-premises host |
| `provision-pipeline-release.sh` | Terminal operator or automation | Creates/reuses a Pipeline and Release definition using REST and Basic PAT authentication |

## Extension contribution lifecycle

The manifest contributes an `ms.vss-web.action` with registered object ID
`generate-pipeline-action`, an invisible `ms.vss-web.control` named
`pipeline-generator-dialog`, and an `ms.vss-web.hub` named
`pipeline-generator-hub`. The action targets several legacy and current
branch-menu contribution IDs because Azure DevOps Server versions expose
different menu surfaces. Both hosted-content contributions map to
`dist/index.html`; the Hub targets `ms.vss-code-web.code-hub-group` and is the
in-host compatibility route when the Server lacks the custom Dialog service.

When Azure DevOps loads the action:

1. `menu-action.js` looks for an ambient VSS SDK.
2. If the ambient SDK is incomplete, it loads the bundled minified SDK and then
   the bundled non-minified file as a fallback.
3. It calls `VSS.init({ usePlatformScripts: true, explicitNotifyLoaded: true })`
   and waits for `VSS.ready`.
4. It preloads and warms the generator assets.
5. It registers an object whose `execute(context)` method calls
   `openGenerator`.
6. `openGenerator` requests the core host page-layout service and calls
   `openCustomDialog` with the fully-qualified control contribution ID.
7. If that service is unavailable, it uses the legacy host navigation service
   to navigate the current Azure DevOps page to the fully-qualified Hub route.
   It never calls `window.open` or obtains an access token.
8. If registration runs before the SDK exposes `register`, it retries every 50
   milliseconds until registration succeeds.

Only bundled SDK assets are loaded. This is intentional: some on-premises hosts
challenge platform SDK asset requests using browser-level Basic authentication,
which can cause repeated login prompts.

## Branch context extraction

Azure DevOps branch menu payloads differ by server version and menu surface.
`menu-action.js` therefore checks multiple shapes for each value.

The action resolves:

- project from the action, repository, or web context;
- repository from `gitRepository`, `repository`, nested item/branch/ref fields,
  or web context;
- repository name from context and finally the `/_git/<name>` URL segment;
- branch from nested branch/ref/item fields, the `version=GB...` query value,
  repository default branch, or the literal fallback `Unknown branch`.

Branch normalization removes a leading `GB` version prefix and
`refs/heads/`. Repository names are URL-decoded and trimmed.

The action builds one set of non-secret hosted-context values. It passes them
as Dialog configuration or as the parent Hub route's query string:

```text
branch=<source-branch>
  &projectId=<project-guid>
  &projectName=<project-name>
  &repoId=<source-repository-guid>
  &repoName=<source-repository-name>
  &hostUri=<collection-uri>
```

## Hosted UI, context, and token acquisition

The preferred path opens a modal Azure DevOps host dialog through service ID
`ms.vss-features.host-page-layout-service`. The action constructs the fully
qualified content ID:

```text
<publisher>.<extension-id>.pipeline-generator-dialog
```

It passes this non-secret configuration to `openCustomDialog`:

```js
{
  branch,
  projectId,
  projectName,
  repoId,
  repoName,
  hostUri
}
```

Inside the host iframe, `ui.js` initializes the SDK, waits for `VSS.ready`,
reads this object with `VSS.getConfiguration()`, and calls
`VSS.getAccessToken()` itself. Consequently the token is issued in the trusted
hosted-content lifecycle for the signed-in user. It never appears in dialog
configuration, a URL, or a cross-page message on the normal path.
Iframe detection does not require `document.referrer`, because some host
referrer policies remove it; successful VSS parent-channel handshake is the
actual trust boundary.

For compatibility with a host that lacks `openCustomDialog`, the manifest also
registers a project-level Azure Repos Hub. The action constructs this parent
route and calls the legacy host navigation service's `navigate` method:

```text
<collection>/<project>/_apps/hub/
  <publisher>.<extension-id>.pipeline-generator-hub
  ?branch=...&projectId=...&projectName=...&repoId=...&repoName=...&hostUri=...
```

The Hub is rendered as an Azure DevOps iframe. `ui.js` obtains the parent query
values through `IHostNavigationService.getCurrentState()` on legacy hosts or
`getQueryParams()` on modern hosts, then performs its own
`VSS.getAccessToken()` call. There is no detached-window, `postMessage`, or
action-token path.

If host token acquisition fails, the generator reveals a **Sign out and
authenticate again** action and an **Open extension authorization** action.
`HostAuthorizationNotFound` is treated as missing collection-level extension
authorization, not merely an expired user session. The authorization action
navigates the parent to `_settings/extensions?tab=installed`, where a Collection
Administrator must select Pipeline Generator and approve its requested scopes.
If no approval action is present, reinstalling the same published version is
the documented recovery for a stale authorization record.

The sign-out action discards its in-memory token and uses
the legacy Azure DevOps host navigation service to navigate the parent page to
the collection-relative `_signout` route. If that service is unavailable, it
navigates the top-level browser window directly. This ends the shared Azure
DevOps browser session rather than merely reloading the extension iframe. The
user completes the server's full login flow and then reopens Generate pipeline
from the target branch. The browser extension does not request or accept a PAT.

Opening `dist/index.html` directly is an offline/error-display mode. It cannot
provision resources because it has neither trusted project context nor a VSS
token.

The hosted page deliberately does not call `VSS.resize()`. On this legacy Azure
DevOps Server, a parameterless call uses `body.scrollWidth` as the requested
contribution width and can create a feedback loop that repeatedly narrows the
iframe. Instead, `html` and `body` are bounded to the host viewport and keep
root overflow hidden. The fixed, full-viewport `.wrapper` is the explicit
vertical scroll container. This avoids the legacy iframe's special root/body
scroll behavior while keeping the form width stable and every control
reachable without resizing the host contribution.

## Runtime state

`dist/ui.js` keeps one in-memory state object:

| Field | Meaning |
| --- | --- |
| `sdk` | Normalized VSS SDK instance when initialized in-frame |
| `accessToken` / `accessTokenError` | Short-lived host token or acquisition error; never persisted |
| `hostUri` | Normalized collection base URI ending in `/` |
| `projectId` | Current Azure DevOps project GUID |
| `rawProjectName` / `projectName` | Display name used for resource names and routes |
| `repoId` | Source repository ID; remains stable across submits and retries |
| `rawRepositoryName` / `repositoryName` | Source repository name; remains stable across submits and retries |
| `generatedRepoId` / `generatedRepositoryName` | Generated repository identity, stored separately so retries cannot overwrite source identity |
| `deploymentTargets` / `deploymentTargetsReady` | Runtime YAML environments plus direct Komodo enabled-server names, and the gate that keeps Submit disabled until both sources are valid |
| `sourceBranch` | Branch selected by the user; used inside generated YAML and filename |
| `branch` | Generated repository branch; fixed to `main` |

`repoId` and the source repository name remain source identity throughout a
submit/retry cycle; generated repository metadata is never written back into
those fields. State lasts only for the lifetime of the page. Refreshing or closing the page
discards the credential and all state.

## Form hydration

The form starts with these defaults:

| Field | Default or derivation |
| --- | --- |
| Pool | `PublishDockerAgent`; merged with project agent queues |
| Service | Lowercase source repository suffix after removing a matching project-name prefix and separator; remains user-editable |
| Environment | Name/domain records loaded from `SharedTemplates/SharedTemplates:/pipeline-generator.yml@main`; `demo` is preferred when present, then inferred from source branch when possible |
| Dockerfile directory | `**`, then first recursively discovered Dockerfile directory |
| Registry address | `registry.buluttakin.com` |
| Registry service | `BulutReg`; merged with Docker Registry service endpoints |
| Komodo server | Loaded directly from Komodo using the central SharedTemplates credential file; only resources with `config.enabled === true` are retained, then the selected environment is used for label inference |
| Target repository | Read-only `<ProjectName>_Azure_DevOps` |

The hosted UI concurrently reads the environment and credential files with the
signed-in user's Bearer token through the Git Items API. It then sends the
credential only in `X-Api-Key` and `X-Api-Secret` headers to Komodo `/read`.
The YAML `environments` list must contain a valid domain for every name, and
the filtered Komodo result must be non-empty. The preferred record shape is
`- name: dev` followed by `domain: bulutdev.ir`; compact
`"dev:bulutdev.ir"` values remain accepted for migration.
A file-read, permission, API, CORS, TLS, or validation failure keeps both
selects and Submit disabled; the extension never restores compiled-in targets.

Environment inference follows this order:

1. A branch containing `master` or `main` maps to `pro`.
2. Otherwise the first configured environment name found as a substring of the
   branch is used.
3. Otherwise the default `demo` remains selected.

Changing the environment selects the first server whose normalized label starts
with that environment. Common abbreviations such as `dev`/`development` and
`pro`/`production` are recognized. A user can then choose another available
target manually. An unmatched environment clears the server selection so the
required field forces an explicit choice.

Agent queue and registry discovery failures are non-fatal: the UI falls back to
the built-in options. Dockerfile discovery failures set `**` and ask the user
to provide a directory.

Both host-dialog and Azure Repos Hub initialization scan the selected source
repository and source branch for Dockerfiles. Discovery affects only the
suggested form value; the user can always enter a path manually.

## Naming and identity

### Generated repository

```text
<projectName>_Azure_DevOps
```

Project casing is preserved. Repository reuse uses exact name equality.

### Support repositories and folders

Step 1 also creates or reuses two project support repositories. Whitespace is
removed from the project name while casing is preserved in repository names:

```text
<ProjectNameWithoutSpaces>_Docker_DevOps
<ProjectNameWithoutSpaces>_Nginx_DevOps
```

Each repository uses `main` and receives `/environments` with
`mattermost_channel=changeme` only when that file is missing. The selected
environment receives starter files instead of empty placeholders:

```text
Docker: /<environment>_<lowercase-project-without-spaces>/compose.yml
Nginx:  /<environment>/<lowercase-project>-<environment>.conf
```

The Compose service/container name is
`<lowercase-project>_<service>_<environment>`. UI/frontend services expose port
80; other services expose port 8080. The Nginx host is
`<lowercase-sanitized-project>.<environment-domain>`. UI/frontend services own
`/`; every other service owns `/<service>/`. Every managed Location configures
Docker DNS with `resolver 127.0.0.11 ipv6=off` and stores the container hostname
in `$target`. Root uses `proxy_pass http://$target:80/`. A non-root route adds
`proxy_pass http://$target:8080` without a URI slash and without `rewrite`, so
the original request URI—including its service prefix—is forwarded unchanged.
This keeps container DNS dynamic. The root Location is
always placed after every other managed Location. WebSocket forwarding is enabled,
`client_max_body_size` is zero, and certificate filenames use the first domain
label. A later run reads the shared Nginx file, identifies its unique HTTPS
`server` by exact `server_name` plus port 443, and enumerates direct-child
Locations with a quote/comment/brace-aware tokenizer. Existing non-root
`/<service>` paths are migrated to `/<service>/`, exact rewrite lines from the
older generated format are removed, direct-host proxy targets become `$target`,
and root is moved below all other generated route blocks. Root retains its
trailing proxy URI slash; non-root proxy targets do not have one. A missing
route is inserted inside managed-route
markers; manual content outside and inside existing Location blocks is
preserved. Missing/malformed markers, unmatched braces, or duplicate matching
HTTPS server blocks stop the edit instead of guessing. Repeated runs converge.

### YAML filename

```text
<sanitized-project>-<sanitized-source-repository>-<SanitizedBranch>To<UPPERCASE-ENVIRONMENT>.yml
```

Each segment is trimmed and lowercased. Slash and backslash runs become `-`;
characters outside word characters, dot, and hyphen become `-`; repeated and
edge hyphens are removed. JavaScript `\w` preserves ASCII letters, digits, and
underscore.

Example:

```text
Project: RideSharing
Repository: RideSharing_Backend
Environment: demo
Branch: feature/defineZones

/ridesharing-ridesharing_backend-Feature-DefineZonesToDEMO.yml
```

### Pipeline and Release names

```text
Pipeline: <generated-yaml-filename>
Release:  <UPPERCASE-SERVICE> <UPPERCASE-ENVIRONMENT>
```

The Pipeline name is exactly the filename returned by
`buildPipelineFilename`, including `.yml` and excluding only the leading
repository path slash. For example, the YAML path
`/ridesharing-ridesharing_backend-Feature-DefineZonesToDEMO.yml` maps to
Pipeline name `ridesharing-ridesharing_backend-Feature-DefineZonesToDEMO.yml`.

For example, Service `api` and Environment `dev` produce Release name
`API DEV`. Environment is mandatory in the Pipeline filename and is expressed
as the destination of the source Branch: `<Branch>To<ENVIRONMENT>`. For example,
Branch `Production` and Environment `soc` produce `ProductionToSOC`. The same
repository and branch can therefore have distinct destination Pipelines.
Release lookup still falls back to Pipeline artifact ID, so a legacy
filename-based Release is renamed and reconciled in place rather than
duplicated.

## Generated YAML contract

`buildPipelineYaml` creates a YAML document with this logical structure:

```yaml
trigger: none

resources:
  repositories:
    - repository: SharedTemplatesRepo
      type: git
      endpoint: ShonizCollection
      name: SharedTemplates/SharedTemplates
      ref: main

    - repository: otherRepo
      type: git
      name: "<ProjectName>/<SourceRepositoryName>"
      ref: refs/heads/<SourceBranch>
      trigger:
        branches:
          include:
            - <SourceBranch>

variables:
- group: KomodoAPI

stages:
- template: build-push-komodo.yml@SharedTemplatesRepo
  parameters:
    pool: '<selected pool>'
    service: '<service>'
    environment: '<environment>'
    dockerfileDir: '<directory or **>'
    repositoryAddress: '<registry host>'
    containerRegistryService: '<service connection>'
    tag: '1.0.$(Build.BuildId)'
    komodoServer: '<selected Komodo server>'
    komodoApiKey: '$(KOMODO_API_KEY)'
    komodoApiSecret: '$(KOMODO_API_SECRET)'
    sourceRepo: otherRepo
```

`trigger: none` disables a trigger for the generated repository itself. The
`otherRepo` repository resource contains the selected source-branch trigger.
The shared template and variable group must already exist and be authorized for
the generated Pipeline.

Form values are currently interpolated directly into YAML strings. Values that
contain a single quote, newline, or YAML control syntax are not escaped. The
current select inputs constrain most values, but `service`, Dockerfile path,
and registry address are free text. Any future expansion to untrusted inputs
must add a YAML-safe serializer or explicit validation.

## Five-step provisioning transaction

Provisioning is a sequential workflow, not an atomic distributed transaction.
Each step updates the status element and attaches its label to any thrown error.

### Step 1: ensure generated and support repositories

The UI lists project repositories using Git API 6.0 and compares exact names.
If the generated, Docker DevOps, or Nginx DevOps repository is missing, it
creates it in the current project. The two support repositories are initialized
idempotently with their root `environments` file, selected-environment starter
configuration, and `main` default branch before Step 2 begins.

### Step 2: add or edit YAML

The UI reads `refs/heads/main` to obtain its current object ID. Missing branches
use Azure DevOps's all-zero object ID. If `main` already exists, it reads the
target path as text. When the existing content is byte-for-byte equal to the
generated YAML, the step returns without a Git Push. Otherwise it chooses Git
change type `add` or `edit`.

For changed or missing content, it submits one Git push containing one ref
update and one commit. The ref's old object ID provides optimistic concurrency.
A concurrent update can therefore cause the push to fail rather than overwrite
an unseen commit.

### Step 3: set default branch

The generated repository is patched to use `refs/heads/main` as its default
branch. This is performed even when the repository already existed.

### Step 4: upsert or migrate Pipeline

The UI searches Pipelines by the desired exact `BranchToEnvironment` name,
then by the 0.1.37 Environment-first name, and finally by the earlier
branch-only filename.

- If an exact filename-named Pipeline exists, it reads the canonical complete
  Build Definition and compares the binding. The Pipelines by-ID response is
  not used for this decision because the target Server omits
  `repository.defaultBranch` from that sparse model.
- If none of those names is present, it searches Build Definitions by the new
  transition YAML path, the 0.1.37 Environment-first path, and the earlier
  branch-only path, in that order. This migrates the existing Pipeline ID
  instead of creating a duplicate. If an old bug created more than one path
  match, the lowest/oldest definition ID is selected deterministically and
  unrelated duplicates are left untouched.
- A correct binding is reused without a write. Folder casing is normalized for
  comparison, so server normalization from `\komodo` to `\KOMODO` is harmless.
- A legacy name or incorrect binding is reconciled through the Build
  Definitions API: GET the complete current definition, preserve its revision,
  modify name/path/YAML/repository/default branch, then PUT it back.
- If no same-name or same-file definition exists, it POSTs a new Pipeline with
  the generated repository ID and exact YAML path.
- Create and update responses must contain a definition ID.

The create URL includes `repositoryId=<generated-repository-id>` in addition to
the repository object in the JSON body. This is required for reliable binding
on the target Azure DevOps Server.

The code never sends `PUT /_apis/pipelines/{id}` because the target server
returns HTTP 405 for that method. Pipeline updates use
`PUT /_apis/build/definitions/{id}?api-version=7.1`, including the latest
revision required by Azure DevOps.

### Step 5: ensure or migrate classic Release definition

The UI reads `PipelineGeneratorReleaseConfig`, derives the Release name, and
searches definitions by exact name.

- If Release creation is disabled, the step returns a visible skipped result.
- It first searches by the desired exact Release name.
- If that name is absent, it searches expanded artifact data for a legacy
  Release whose Build artifact references the same Pipeline ID.
- It reads any match in full and compares name, folder, artifact/repository,
  environment, queue, Bash task/script, event condition, and automated
  approvals. A matching definition is reused; a mismatch is updated with its
  existing ID, revision, and environment identity preserved.
- If no matching name or Pipeline artifact exists, it POSTs a new definition.
- A duplicate-name conflict is re-read and reconciled instead of immediately
  failing.

In parallel with resolving the selected queue and Bash source, the UI resolves
the exact project Variable Group `KomodoAPI` using `actionFilter=Use`. It fails
before a Release write if the group is unavailable or does not declare all of
`AZP_TOKEN`, `KOMODO_API_KEY`, and `KOMODO_API_SECRET`. Only names and secret
metadata are inspected; values are never written to logs or persisted by the
extension. The resolved numeric ID is linked through the Release definition's
top-level `variableGroups` array. Environment-level `variableGroups` remains
empty, matching the target Server's working classic Release shape. Existing
additional definition-level Variable Groups are preserved during reconciliation.

The created definition contains one primary Build artifact pointing to the
Pipeline ID, one environment, and one agent-based deployment phase containing
an inline Bash v3 task. By default `release-config.js` points to packaged asset
`release-inline-task.sh`; the UI fetches it at definition-generation time and
stores its complete text in `workflowTasks[0].inputs.script`. The source can
also be literal inline content or a same-collection Azure Repos file. In every
mode the resulting Release task is Inline, not a file-path task.

The default wrapper expands the three secret Release variables at execution
time, clones `SharedTemplates/release-komodo.sh`, falls back to the Git Items
REST API for a single-file download, and executes the downloaded script. Shell
xtrace is deliberately disabled around the secret-derived Authorization
header, and Git/curl/wget receive it through environment/config channels rather
than process arguments. Xtrace begins only for the final shared-script command.

The environment has a `ReleaseStarted` event condition, automated pre- and
post-deployment approvals, 30-day/3-release retention, and the selected queue.
No continuous deployment trigger is configured.

## Reconciliation and retry behavior

| Resource | Lookup identity | Existing-resource behavior |
| --- | --- | --- |
| Generated/support repository | Exact repository name | Reuse; add missing bootstrap files and merge only missing Nginx service Locations |
| YAML file | Generated path on `main` | Reuse without Push when byte-identical; otherwise add/edit with a new commit |
| Default branch | Repository ID | Always patch to `refs/heads/main` |
| Pipeline | Exact BranchToEnvironment name/path, then 0.1.37 Environment-first and older branch-only identities | Reuse or GET-modify-PUT through Build Definitions |
| Release definition | Exact Release name, then Pipeline artifact ID | Reuse or reconcile through Release Definitions PUT |

Because there is no rollback, a later failure leaves earlier successful
resources in place. This is intentional and makes most retries convergent. For
example, if Release creation fails, rerunning reuses unchanged YAML, reuses or
repairs the Pipeline, and retries Release creation.

Release configuration changes now propagate to the first definition matching
the exact desired name or Pipeline artifact ID. Unrelated historical duplicates
are not deleted automatically.

## Completion and error behavior

When all enabled operations succeed, the UI displays the Pipeline ID and
Release definition result and stays on the form. It renders three explicit
links:

```text
<Nginx repository>/<environment>/<project>-<environment>.conf
<Docker repository>/<environment>_<project>/compose.yml
<collection>/<project>/_build?definitionId=<pipeline-id>
```

The first two links let the operator review/edit the generated starters before
opening and manually running the Pipeline. Successful provisioning performs no
automatic navigation and queues no Pipeline run.

HTTP error bodies are converted to text, stripped of HTML/script/style markup,
collapsed to one line, and truncated to 500 characters. Errors are tagged as
Pipeline or Release domain errors so the UI can display the relevant permission
hint. HTTP 401, 403, TF400813, and matching authorization messages are rendered
as access-denied failures and reveal the full-session reauthentication action.

The request header `X-TFS-FedAuthRedirect: Suppress` asks Azure DevOps to return
an API error instead of redirecting the extension iframe to an interactive
login page.

## Architectural constraints

- The generated repository and `main` branch are hard-coded conventions.
- Shared template repository, environment/credential file paths, Variable
  Group, template filename, Pipeline folder, and API versions are constants in
  `dist/ui.js`.
- Name-based lookups use exact equality and return the first match.
- List calls do not follow continuation tokens or implement pagination. Large
  projects can hide a repository, Pipeline, Release, queue, or service endpoint
  beyond the first response page.
- Pipeline names intentionally include the Branch-to-Environment transition YAML filename;
  Release names intentionally contain only Service and Environment.
- Pipeline and Release folder comparison is case-insensitive.
- Pipeline migration depends on Build Definitions list filtering by repository
  and the desired or legacy YAML filename; Release migration filters expanded artifacts by type
  `Build` and source ID `<projectId>:<pipelineId>`.
- There is no transaction or automatic cleanup for partial failure.
- There is no YAML parser/serializer; generated text is assembled manually.
- Client-side code is shipped unobfuscated and must not contain secrets.
- The central Server-Read credential is intentionally browser-readable by
  operator policy. It must stay outside extension assets and must not be logged,
  persisted to browser storage, or reused for write-capable Komodo access.
- The browser runtime accepts only the short-lived host token. PAT support is
  limited to the separate terminal provisioner and is never exposed in the UI.
- Although the manifest declares an Azure DevOps Services target, Release REST
  calls are built from the collection host and do not separately resolve the
  Services `vsrm.dev.azure.com` host. The full workflow is currently verified
  only on the documented on-premises server.
- An existing shared target repository is trusted by name. Project repository
  permissions are the security boundary.
- Pipeline or Release execution is outside this extension's workflow. The
  extension creates definitions only.
