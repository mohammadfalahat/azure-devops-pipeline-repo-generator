# Architecture and runtime flow

This document describes version 0.1.70 from the implementation in
`vss-extension.json`, `dist/menu-action.js`, `dist/ui.js`, and
`dist/release-config.js`.

## System boundary

Pipeline Generator is a client-only Azure DevOps web extension. It has:

- no application server;
- no database;
- no persisted runtime credential;
- no build-time generated JavaScript;
- no background worker or webhook dependency.

All Azure DevOps provisioning calls are made directly from the user's browser.
Project resources are read and written in the collection that hosts the
extension, while the two central configuration files are always read from the
sibling collection `ShonizCollection`. The security principal is the signed-in
user represented by the token returned by `VSS.getAccessToken()`.
The manifest includes separate read scopes for agent pools/queues, service
endpoints, and Variable Groups; Build and Release scopes do not implicitly
grant those resource-area permissions. The browser also reads the central
`/komodo-servers-creds.env` file from
`ShonizCollection/SharedTemplates/SharedTemplates`, then calls the Komodo
1.19.x read API using the operator-designated non-confidential Server-Read
credential. Credential values remain in page/request memory only and are never
logged or persisted. Komodo must allow the Azure DevOps origin through CORS.

The separate shell script is an operational fallback, not a backend for the
extension. It receives a PAT and assumes the target repository and YAML file
already exist.

## Component model

| Component | Loaded by | Main responsibilities |
| --- | --- | --- |
| `vss-extension.json` | Azure DevOps extension host | Declares the normal and Monorepo branch-menu actions, Dialog control, Azure Repos Hub, supported hosts, addressable files, and token scopes |
| `menu-action.html` / `menu-action.js` | Hidden action contribution iframe | Initializes VSS SDK, registers `generate-pipeline-action` and `generate-monorepo-action`, extracts branch context/mode, warms assets, and asks the host to open the form |
| `index.html` / `ui.js` | Dialog or `pipeline-generator-hub` host iframe | Reads host configuration/navigation state, obtains the current-user host token, hydrates the form, performs five provisioning steps, displays errors, and renders Nginx/Compose/Pipeline review links |
| `vendor/tom-select/*` | Loaded by `index.html` before `ui.js` | Provides the locally packaged Tom Select 2.6.2 JavaScript/CSS editable combobox implementation and its Apache-2.0 license |
| `release-config.js` | Loaded before `ui.js` | Exposes immutable `window.PipelineGeneratorReleaseConfig` with Release settings, `KomodoAPI` requirements, and Bash source selection |
| `release-inline-task.sh` | Fetched by `ui.js` from the installed extension assets | Provides the wrapper text embedded into the classic Release Bash task |
| `monorepo-build.cjs` | Maintained mirror of `SharedTemplates:/monorepo/mr-build.cjs` | Discovers buildable Nx apps, computes affected apps, applies shell rebuild-all, resolves output paths, and creates `mr-drop` |
| `monorepo-release-inline-task.sh` | Embedded by `ui.js` into MR classic Releases | Commits manifest image tags to the shared Compose source, deploys the Git-linked Komodo Stack, polls its Update, and commits rollback tags on failure |
| `VSS.SDK*.js` | Action and generator pages | Supplies the legacy VSS extension APIs required by the supported on-premises host |
| `provision-pipeline-release.sh` | Terminal operator or automation | Creates/reuses a Pipeline and Release definition using REST and Basic PAT authentication |

## Extension contribution lifecycle

The manifest contributes two `ms.vss-web.action` objects with registered object
IDs `generate-pipeline-action` and `generate-monorepo-action`, an invisible `ms.vss-web.control` named
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
5. It registers both action objects; their `execute(context)` methods call
   `openGenerator` with `pipeline` or `monorepo` mode.
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
| Service | Lowercase source repository suffix after removing a matching project-name prefix and separator; whitespace is replaced only with `_` (other punctuation is preserved) and the value remains user-editable |
| Environment | Editable Tom Select 2.6.2 combobox with name/domain suggestions (plus legacy `projects_root` metadata) loaded from `ShonizCollection/SharedTemplates/SharedTemplates:/pipeline-generator.yml@main`; `demo` is preferred when present, then inferred from source branch when possible |
| Stack | Editable Tom Select 2.6.2 combobox with `default` selected; suggestions are discovered from existing top-level Docker DevOps Compose directories, while free text creates another isolated Stack |
| Dockerfile directory | `**`, then first recursively discovered Dockerfile directory |
| Registry address | `registry.buluttakin.com` |
| Registry service | `BulutReg`; merged with Docker Registry service endpoints |
| Komodo server | Loaded directly from Komodo using the central SharedTemplates credential file; only resources with `config.enabled === true` are retained, then the selected environment is used for label inference |
| Target repository | Read-only `<ProjectName>_Azure_DevOps` |

The hosted UI concurrently reads the environment and credential files through
the same-origin signed-in browser session, without forwarding the current
collection's Bearer token to the sibling collection. Current-project REST calls
continue to use that scoped Bearer token. The UI then sends the
credential only in `X-Api-Key` and `X-Api-Secret` headers to Komodo `/read`.
The YAML `environments` list must contain a valid domain for every name, and
the filtered Komodo result must be non-empty. The preferred record shape is
`- name: dev` followed by `domain: bulutdev.ir`; compact
`"dev:bulutdev.ir"` values remain accepted for migration.
The user may type a value outside that configured list. A case-insensitive
configured-name match uses its paired domain and enables Nginx generation; an
unlisted value remains the Pipeline/Release/Compose Environment but creates or
updates no Nginx repository or configuration.
The normalized Stack defaults to `default`. That value preserves every legacy
identity. A non-default value is inserted into Compose/Nginx directories,
Pipeline/Release names, image/container names, and Komodo Repo/Stack resources.
Both comboboxes use locally packaged JavaScript/CSS, open their full option
list on click or focus, retain searchable keyboard navigation, and remain
editable after values are discovered.
A file-read, permission, API, CORS, TLS, or validation failure keeps the
Environment combobox, Komodo Server select, and Submit disabled; the
extension never restores compiled-in targets.

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

Immediately before Step 1 performs any repository write, the generator lists
the current project's repositories and builds a deterministic ownership plan
for the two shared base Locations. A repository whose name equals the project
name case-insensitively is forced to frontend and has highest priority for `/`.
Otherwise the generic semantic-specificity score prefers a pure recognized role
name over a composite name, then fewer non-role qualifiers, an explicit token
boundary over a joined prefix/suffix, shorter qualifier text, fewer tokens, and
finally lexical order. Recognized role aliases are equivalent at the pure-name
level; there is no table of preferred repository-name pairs. The same ranking
applies independently to `/api/`. Thus `front` owns `/` ahead of `front_panel`,
`frontend` owns it ahead of `frontend_dashboard`, and `api` owns `/api/` ahead
of `api_admin`, regardless of which Pipeline is generated first. Joined forms
such as `frontpanel` and `apiadmin` are also classified. A lower-priority
candidate keeps its frontend/backend port but falls back to its own normalized
path, such as `/front-panel/` or `/api-admin/`.

## Naming and identity

### Generated repository

```text
<projectName>_Azure_DevOps
```

Project casing is preserved. Repository reuse uses exact name equality.

### Support repositories and folders

Step 1 creates or reuses the Docker support repository and, only for a
configured Environment, the Nginx support repository. Whitespace is removed
from the project name while casing is preserved in repository names:

```text
<ProjectNameWithoutSpaces>_Docker_DevOps
<ProjectNameWithoutSpaces>_Nginx_DevOps
```

Each created repository uses `main`. Compose is generated for every valid
Environment value; the Nginx path is generated only when the value matches a
configured Environment:

```text
Docker default: /<environment>_<lowercase-project-without-spaces>/compose.yml
Docker custom:  /<environment>_<stack>_<lowercase-project-without-spaces>/compose.yml
Nginx default:  /<environment>/<lowercase-project>-<environment>.conf
Nginx custom:   /<environment>_<stack>/<lowercase-project>-<environment>.conf
```

The Compose service/container name is
`<lowercase-project>_<service>[_<stack>]_<environment>`; the Stack segment is
omitted for `default`. Images use the analogous
`<service>[-<stack>]-<environment>` suffix. UI/frontend services expose port
80; backend and other services expose port 8080. The Nginx host is
`<lowercase-sanitized-project>.<environment-domain>`. Routing uses semantic
tokens rather than a short exact-name list. `ui`, `front`, `frontend`, `web`,
`fe`, `website`, `client`, `portal`, and `spa` indicate frontend; `api`, `back`,
`backend`, `be`, `server`, `bff`, `rest`, `graphql`, and `gateway` indicate backend.
Unversioned frontend owns `/`, unversioned backend owns `/api/`, and every
other service owns `/<service>/`. Recognized newer markers include `new`,
`refactor`, `rewrite`, `revamp`, `next`, `nextgen`, `modern`, `latest`, and
numeric forms such as `v2`, `version2`, `ver2`, or `r2`. A frontend variant
uses `/<variant>/`; a backend variant uses `/api/<variant>/`. For example,
`UI_V2` maps to `/v2/`, `BACK_v2` maps to `/api/v2/`, `NewUI` maps to `/new/`,
and `api_refactor` maps to `/api/refactor/`. Numeric markers are canonicalized
to lowercase `vN`; if several are present, the highest is used. Every managed Location configures
Docker DNS with `resolver 127.0.0.11 ipv6=off` and stores the container hostname
in `$target`. Root uses `proxy_pass http://$target:80`. A non-root route adds
`proxy_pass http://$target:8080` without a URI slash and without `rewrite`, so
the original request URI—including its service prefix—is forwarded unchanged.
This keeps container DNS dynamic. The root Location is
always placed after every other managed Location. WebSocket forwarding is enabled,
`client_max_body_size` is zero, and certificate filenames use the complete
environment domain (for example, `bulutco.cloud.pem` and `bulutco.cloud.key`).
A later run reads the shared Nginx file, identifies its unique HTTPS
`server` by exact `server_name` plus port 443, and enumerates direct-child
Locations with a quote/comment/brace-aware tokenizer. Existing non-root
legacy service-name paths are migrated to their semantic canonical path, exact
rewrite lines from the older generated format are removed, direct-host proxy targets become `$target`,
exact legacy generated certificate paths based on only the domain's first
label are migrated to the complete Environment domain, and root is moved below
all other generated route blocks. Certificate migration is scoped to the
matching HTTPS server and does not alter custom paths. Neither root nor
non-root proxy targets have a URI slash. A missing
route is inserted inside managed-route
markers; manual content outside and inside existing Location blocks is
preserved. Missing/malformed markers, unmatched braces, or duplicate matching
HTTPS server blocks stop the edit instead of guessing. Repeated runs converge.
Repository-priority overrides are applied while normalizing managed blocks, so
a lower-priority repository cannot retain or later reclaim `/` or `/api/`.

### YAML filename

```text
<sanitized-project>-<sanitized-source-repository>[-MR]-<sanitized-service>[-<stack>]-<SanitizedBranch>To<UPPERCASE-ENVIRONMENT>.yml
```

Project, repository, and Service segments are trimmed and lowercased. Slash and
backslash runs become `-`; characters outside word characters, dot, and hyphen
become `-`; repeated and edge hyphens are removed. The Branch retains its
word-leading capitalization and the Environment is uppercased. JavaScript `\w`
preserves ASCII letters, digits, and underscore.

Example:

```text
Project: RideSharing
Repository: RideSharing_Backend
Service: api
Environment: demo
Branch: feature/defineZones

/ridesharing-ridesharing_backend-api-Feature-DefineZonesToDEMO.yml
```

### Pipeline and Release names

```text
Pipeline:   <project>-<repository>[-MR]-<service>[-<stack>]-<Branch>To<ENVIRONMENT>.yml
Release:    <UPPERCASE-SERVICE> [<UPPERCASE-STACK>] <UPPERCASE-ENVIRONMENT>
MR Release: MR <UPPERCASE-SERVICE> [<UPPERCASE-STACK>] <UPPERCASE-ENVIRONMENT>
```

The Pipeline name is exactly the filename returned by
`buildPipelineFilename`, including `.yml` and excluding only the leading
repository path slash. For example, the YAML path
`/ridesharing-ridesharing_backend-api-Feature-DefineZonesToDEMO.yml` maps to
Pipeline name `ridesharing-ridesharing_backend-api-Feature-DefineZonesToDEMO.yml`.

For example, Service `api` and Environment `dev` produce Release name
`API DEV`. Service and Environment are mandatory in the Pipeline filename;
Environment is expressed as the destination of the source Branch:
`<Branch>To<ENVIRONMENT>`. For example, Branch `Production` and Environment
`soc` produce `ProductionToSOC`. The same
repository and branch can therefore have distinct Service and destination
Pipelines without sharing a YAML path or Build Definition identity.
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
    stack: '<stack; default when omitted>'
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
If the generated or Docker DevOps repository is missing, it creates it in the
current project. For a configured Environment it does the same for Nginx
DevOps; for a custom Environment it makes no Nginx read or write. Created
support repositories are initialized idempotently with their
selected-environment starter configuration and `main` default branch before
Step 2 begins. The extension does not create a root `environments` file.

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

The UI searches Pipelines by the desired exact Service-aware
`BranchToEnvironment` name, then by the immediate predecessor's Service-less
transition name, the 0.1.37 Environment-first name, and finally by the earlier
branch-only filename.

- If an exact filename-named Pipeline exists, it reads the canonical complete
  Build Definition and compares the binding. The Pipelines by-ID response is
  not used for this decision because the target Server omits
  `repository.defaultBranch` from that sparse model.
- If none of those names is present, it searches Build Definitions by the new
  Service-aware transition YAML path, the immediately preceding Service-less
  transition path, the 0.1.37 Environment-first path, and the earlier
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

## Nx Monorepo (`MR`) execution model

Monorepo mode uses the same five provisioning steps but selects separate
renderers and identities. The generated Pipeline/Release live under
`\komodo\MR`, the Pipeline filename contains `-MR-<service>[-<stack>]-<Branch>To<ENV>`,
and the Release is named `MR <SERVICE> [<STACK>] <ENV>`. The immediately preceding
Service-less MR Pipeline is eligible for in-place migration; normal Pipeline
definitions are never considered legacy candidates for MR reconciliation.

Step 1 creates/reuses the same project Docker and Nginx repositories and merges
the Monorepo runtime into the selected project/Environment/Stack Compose. The
default Stack reuses the legacy paths; a custom Stack uses separate Compose and
Nginx directories. The logical service has one immutable Nginx
image for shell/static assets and an optional immutable Node image for BFF output.
Before any repository write, the browser calls Komodo `ListDockerNetworks` for
the selected Server and requires an exact existing `nginx-network` or
`nginx-net`. Compose keeps an existing logical network key when possible and
sets its external `name` to the resolved host network. This lets existing
services continue referring to `nginx-network` while a target host actually
uses `nginx-net`, without renaming unrelated service blocks.
There are no host bind mounts, runtime working-directory overrides, or project
Dockerfiles. The outer Nginx configuration
routes `/bff/` to the BFF and `/` to the static runtime, reserves `/api/` for
the main application backend, uses Docker's dynamic
resolver, performs no rewrite, and keeps `/` after non-root Locations.
The Compose file stays on `main` in the project's Docker DevOps repository and
is the deployment source of truth; no generated or downloaded target-side copy
replaces it.
Reconciliation migrates legacy bare runtime images and removes legacy managed
mount/command fields, while preserving immutable active tags, custom fields,
and every unrelated Compose service. New managed image fields are stable
repository references whose tag comes from service-specific keys in the
adjacent tracked `.env`; an active legacy hard-coded tag is preserved until the
next Release can migrate Compose and `.env` atomically.

Step 2 atomically pushes two files into the generated repository:

- the MR Pipeline YAML, which references
  `monorepo/pipeline.yml@SharedTemplatesRepo` (managed on every generator rerun);
- `/.devops/deployments.yml` (created only when absent, preserving operator
  edits on later runs).

The central template checks out the generated repository, source Monorepo, and
SharedTemplates, logs in through the selected Docker Registry service connection,
then runs `mr-build.cjs`. `package-images.sh` uses the generic SharedTemplates
Dockerfiles to create/push immutable static and optional BFF images.
It also creates or updates the normal Komodo Repo pointing to
`<Project>_Docker_DevOps@main` and partially reconciles the same normal Docker Stack whose
`linked_repo`, `run_directory`, and `file_paths` resolve the exact Git-managed
shared `compose.yml`. Existing Stack environment and extra arguments are retained;
only the BFF profile is reconciled. The Pipeline configures resources but never deploys the Stack.
At Build time the runner installs with `pnpm install --frozen-lockfile`, asks Nx
for buildable applications and affected applications, and invokes the contract
build command separately for each affected project with `{projects}` replaced
by that project. Corepack and pnpm resolve through the SharedTemplates
`npmRegistry` parameter, which defaults to the internal Nexus
`https://registry.buluttakin.com/repository/npm-group`; pnpm's store and
Corepack home are mounted from the self-hosted agent cache. A configured
shell/host application in the affected set
promotes the run to rebuild all applications. An ordinary failed project is
recorded and omitted from the overlay. The prior static image is extracted first,
so its deployed version remains;
successful projects continue and the Build is marked `SucceededWithIssues`.
A failed shell blocks the Build. Without a previous image, ordinary failures
that leave no baseline also block it. Nx metadata supplies output paths. The
`mr-drop` artifact contains schema-2 image references, inventory, successful
outputs, and failed-project metadata; the deployable payload itself is in the registry.

The MR Release embeds `monorepo-release-inline-task.sh`. It reads the downloaded
manifest, clones the Docker DevOps repository, keeps managed Compose image
repositories stable, updates the exact service/BFF tag keys in the adjacent
`.env`, and pushes a release commit. A legacy hard-coded Compose image is
migrated to `${service_key}` form in that same commit. It then calls Komodo `/execute` with `DeployStack`
for the shared Git-linked Docker Stack and polls `/read` with `GetUpdate` until the Update
is `Complete`; `success` must be true.
Static project directories are symlinked below the shell root by Nx project
name, making `/<project-name>/` available through the catch-all static Runtime.
The BFF profile is enabled only when the manifest contains a BFF image. A failed
deployment causes a rollback commit restoring the exact prior Compose/`.env` state and a best-effort
redeploy; the Release remains failed so infrastructure errors are visible.

The browser's central Server-list key remains Server-Read only. The separate
credentials expanded from the current project's `KomodoAPI` Variable Group at
Build/Release execution need Repo/Stack read-create-update permissions,
`DeployStack` execution, Registry push access in Build, and ADO Git read/write.
Terminal permission is not required. Secret headers are fed through config/stdin
or process environment and shell xtrace is not enabled.

Komodo can temporarily reject a Stack mutation with `Stack busy` while another
operation holds the resource. Build-side `CreateStack`/`UpdateStack` and
Release-side `DeployStack` therefore retry only that response for five total
attempts with a five-second interval. `KOMODO_STACK_BUSY_MAX_ATTEMPTS` and
`KOMODO_STACK_BUSY_RETRY_SECONDS` can tune the bounded policy; other errors are
not retried.

## Reconciliation and retry behavior

| Resource | Lookup identity | Existing-resource behavior |
| --- | --- | --- |
| Generated/support repository | Exact repository name | Before writes, resolve the selected Server's exact `nginx-network`/`nginx-net`; reuse repositories, add missing bootstrap files, map the actual external network, and merge only missing Compose services and Nginx Locations |
| YAML file | Generated path on `main` | Reuse without Push when byte-identical; otherwise add/edit with a new commit |
| Default branch | Repository ID | Always patch to `refs/heads/main` |
| Pipeline | Exact Service-aware BranchToEnvironment name/path, then Service-less transition, 0.1.37 Environment-first, and older branch-only identities | Reuse or GET-modify-PUT through Build Definitions |
| Release definition | Exact Release name, then Pipeline artifact ID | Reuse or reconcile through Release Definitions PUT |
| MR deployment contract | `/.devops/deployments.yml` on `main` | Create when missing; preserve all later edits |
| MR Compose Git source | `<Project>_Docker_DevOps:/<environment>[_<stack>]_<project>/{compose.yml,.env}@main` | Reuse the shared Compose; merge absent services, migrate legacy runtime fields, keep stable image repositories in Compose, and store managed immutable tags in `.env` without changing unrelated/operator-edited services |
| MR Komodo Repo/Stack | `<Project>_Docker_DevOps-<environment>[-<stack>]`; the custom suffix is omitted for `default` | Apply only partial Stack updates, preserving unrelated config; deploy only from Release |
| MR runtime state | Immutable Registry tags referenced by managed Compose `.env` keys | Hydrate the prior image, overlay affected outputs, push a new build tag, and restore the exact prior Compose/`.env` state on failed deployment |

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
<Nginx repository>/<environment>[_<stack>]/<project>-<environment>.conf
<Docker repository>/<environment>[_<stack>]_<project>/compose.yml
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
- Pipeline names intentionally include Service and the Branch-to-Environment transition YAML filename;
  normal Release names contain Service and Environment, while MR names also contain the MR marker.
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
