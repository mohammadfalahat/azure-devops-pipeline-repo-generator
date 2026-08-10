# Azure DevOps REST contracts

This document is the integration contract between Pipeline Generator 0.1.32
and Azure DevOps. Paths are relative to the collection base URI unless stated
otherwise.

## Base URI construction

The browser obtains the collection URI from VSS web context when available. As
a fallback, it derives the base from `document.referrer` and supports both
common on-premises layouts:

```text
https://host/<collection>/
https://host/tfs/<collection>/
```

`hostUri` is normalized to remove a trailing `/_apis` and to end with one `/`.
Project-scoped calls then use:

```text
<hostUri><url-encoded-project-id>/_apis/...
```

The shell path constructs URLs from `ADO_URL`, `COLLECTION`, and `PROJECT`. It
does not discover or normalize a collection URL.

The same collection host is also used for classic Release calls. This is the
verified Azure DevOps Server topology. Azure DevOps Services commonly exposes
Release APIs through `vsrm.dev.azure.com`; the current code does not resolve
that host separately, so the manifest's Cloud target must not be interpreted as
proof that the complete Release workflow is Cloud-compatible.

## Authentication and authorization

### Browser extension

The branch action opens the generator through `openCustomDialog` when the host
supports it. Otherwise it navigates the current Azure DevOps page to the
extension's Azure Repos Hub. The resulting hosted iframe completes its own VSS
SDK handshake, reads non-secret branch/repository context from
`VSS.getConfiguration()` or host navigation state, calls
`VSS.getAccessToken()` for the current signed-in user, and sends:

```http
Authorization: Bearer <short-lived-extension-token>
X-TFS-FedAuthRedirect: Suppress
```

The token is opaque and must be treated as an OAuth/session token even when it
does not look like a JWT. It must not be converted to PAT-style Basic
authentication. `X-TFS-FedAuthRedirect: Suppress` prevents an API request from
silently becoming an HTML login redirect.

On both hosted paths, the token exists only in the generator's JavaScript
memory. It is not passed in Dialog configuration or Hub query state and is
never written to local storage, session storage, cookies, URLs, logs,
repository files, or the generated YAML. The branch action never obtains a
token and never opens a detached window.

If Azure DevOps Server cannot issue the host token, the browser UI does not
accept a PAT. For `HostAuthorizationNotFound`, it links the parent host to the
collection-relative `_settings/extensions?tab=installed` page. A Collection
Administrator must select Pipeline Generator and authorize the manifest
scopes. If no authorization action appears, uninstalling and reinstalling the
same published version can recreate a stale authorization record.

The separate full-session recovery action clears the in-memory token and
navigates the parent Azure DevOps host to `_signout` through the host navigation
service, falling back to top-level navigation. Signing in again is useful for a
stale browser session but does not replace collection-level extension scope
authorization.

### Shell provisioner

`scripts/provision-pipeline-release.sh` receives `AZP_TOKEN` and uses curl Basic
authentication. It feeds curl's `user = ":<token>"` setting through curl config
stdin, so the token is not expanded into curl's process arguments. The PAT is
still present in the shell script's environment and must be short-lived,
least-privileged, and unset after use.

### Manifest scopes

| Scope | Why it is requested |
| --- | --- |
| `vso.code` | Read repositories, refs, source tree, YAML/script files |
| `vso.code_manage` | Create the generated repository, push YAML, set default branch |
| `vso.project` | Read current project context/metadata |
| `vso.build` | Read Pipeline and Build definitions |
| `vso.build_execute` | Create or edit Pipeline definitions |
| `vso.release` | Read classic Release definitions |
| `vso.release_manage` | Create classic Release definitions |
| `vso.agentpools` | Read project queues and resolve the selected Release queue ID |
| `vso.serviceendpoint` | Read Docker Registry service-connection names for the form |
| `vso.variablegroups_read` | Resolve `KomodoAPI`, verify its required variable names, and link its numeric ID to the Release definition |

Requested scopes only define what a host token may contain. They do not bypass
Azure DevOps ACLs. The signed-in identity still needs:

- Repos: Read and Contribute; Create repository when the generated repository
  does not exist; permission to edit its default branch.
- Pipelines: View, Create pipeline, and Edit pipeline, including folder
  security for `\komodo` when configured.
- Releases: View releases and Manage release definitions, including Release
  folder security for `\komodo` when configured.
- Agent queue: Use permission on the form's selected queue.
- Variable Groups: permission to use/read `KomodoAPI`; the group must declare
  secret variables `AZP_TOKEN`, `KOMODO_API_KEY`, and `KOMODO_API_SECRET`.
- Service connections and the YAML Pipeline's own Variable Group references:
  authorization required when the generated Pipeline actually runs.
- Optional Release script repository: Read permission on the configured project
  and repository.

Publishing or installing the extension is a separate administrative operation
and requires extension-management permissions that runtime users do not need.

## Browser endpoint matrix

`{project}` means the URL-encoded project GUID. `{repo}` is a repository GUID.
All calls send the Bearer and redirect-suppression headers.

| Operation | Method and path | API version | Expected behavior |
| --- | --- | --- | --- |
| List repositories | `GET {project}/_apis/git/repositories` | `6.0` | Find generated repository by exact name |
| Create repository | `POST {project}/_apis/git/repositories` | `6.0` | Body contains name and current project ID |
| Read branch ref | `GET {project}/_apis/git/repositories/{repo}/refs?filter=heads/main` | `6.0` | Return current object ID or use zero ID when missing |
| Check YAML path | `GET {project}/_apis/git/repositories/{repo}/items?path=...&versionDescriptor...` | `6.0` | HTTP 404 means the file does not exist |
| Read existing YAML text | Same Git Items endpoint with `$format=text` | `6.0` | Equal content means no Push and no commit |
| Push YAML | `POST {project}/_apis/git/repositories/{repo}/pushes` | `6.0` | Add/edit one file through one commit/ref update |
| Set default branch | `PATCH {project}/_apis/git/repositories/{repo}` | `6.0` | Set `defaultBranch` to `refs/heads/main` |
| Scan Dockerfiles | `GET {project}/_apis/git/repositories/{sourceRepo}/items?recursionLevel=Full...` | `6.0` | Find files whose basename is `Dockerfile` |
| Read one repository | `GET {project}/_apis/git/repositories/{repo}` | `6.0` | Fallback source-repository name lookup |
| List agent queues | `GET {project}/_apis/distributedtask/queues` | `6.0` | Populate Pool options and resolve Release queue ID |
| Resolve Release Variable Group | `GET {project}/_apis/distributedtask/variablegroups?groupName=KomodoAPI&actionFilter=Use` | `7.1` | Find the exact group, validate required variable names, and retain only its numeric ID |
| List Docker Registry endpoints | `GET {project}/_apis/serviceendpoint/endpoints?type=dockerregistry&projectIds=...` | `6.0` | Populate registry service options |
| List Pipelines | `GET {project}/_apis/pipelines` | `7.1-preview.1` | Find Pipeline by exact name |
| Read Pipeline | `GET {project}/_apis/pipelines/{id}` | `7.1-preview.1` | Inspect an exact-name Pipeline binding |
| Create Pipeline | `POST {project}/_apis/pipelines?repositoryId={repo}` | `7.1-preview.1` | Create YAML Pipeline and return a non-empty ID |
| Find Build Definition by YAML | `GET {project}/_apis/build/definitions?repositoryId=...&yamlFilename=...&includeAllProperties=true` | `7.1` | Locate a legacy same-file Pipeline when its name differs |
| Read Build Definition | `GET {project}/_apis/build/definitions/{id}` | `7.1` | Obtain the complete definition and current revision |
| Update Build Definition | `PUT {project}/_apis/build/definitions/{id}` | `7.1` | Rename/rebind Pipeline using GET-modify-PUT |
| List Release definitions | `GET {project}/_apis/release/definitions?...` | `7.1-preview.4` | Find by exact name or expanded Build artifact source ID |
| Read Release definition | `GET {project}/_apis/release/definitions/{id}` | `7.1-preview.4` | Obtain complete definition, revision, and environment identity |
| Create Release definition | `POST {project}/_apis/release/definitions` | `7.1-preview.4` | Create artifact/environment/job definition |
| Update Release definition | `PUT {project}/_apis/release/definitions` | `7.1-preview.4` | Reconcile existing same-name/same-Pipeline definition |
| List script repositories | `GET {scriptProject}/_apis/git/repositories` | `6.0` | Resolve optional script repository by ID or name |
| Read Release script | `GET {scriptProject}/_apis/git/repositories/{repo}/items?...&%24format=text` | `6.0` | Return non-empty Bash source as text |

The UI does not call the Pipeline run API or Release instance API.

## Git write contract

For an absent `main` branch the ref update uses Azure DevOps's zero object ID:

```json
{
  "refUpdates": [
    {
      "name": "refs/heads/main",
      "oldObjectId": "0000000000000000000000000000000000000000"
    }
  ],
  "commits": [
    {
      "comment": "Add pipeline generator defaults",
      "changes": [
        {
          "changeType": "add",
          "item": { "path": "/<generated-name>.yml" },
          "newContent": {
            "content": "<generated-yaml>",
            "contentType": "rawtext"
          }
        }
      ]
    }
  ]
}
```

For an existing branch, `oldObjectId` is the current ref object ID. If the file
exists, `changeType` becomes `edit` and the commit message starts with `Update`.
Azure DevOps performs the concurrency check against `oldObjectId`.

The UI currently does not return or persist the push response; downstream
provisioning relies on the POST completing successfully and on the server being
able to resolve the new path immediately.

## Pipeline create and reconciliation contract

The desired Pipeline body is:

```json
{
  "name": "<exact-generated-yaml-filename>.yml",
  "folder": "\\komodo",
  "configuration": {
    "type": "yaml",
    "path": "/<generated-file>.yml",
    "repository": {
      "id": "<generated-repository-guid>",
      "name": "<Project>/<Project>_Azure_DevOps",
      "type": "azureReposGit",
      "defaultBranch": "refs/heads/main"
    }
  }
}
```

On creation, `repositoryId=<generated-repository-guid>` is also required in the
query string. The target on-premises server can otherwise fail to bind the YAML
repository even though the body contains its ID.

The Pipeline display name is the exact generated YAML filename, including its
`.yml` suffix. Before writing, the UI compares:

- `folder` against `\komodo`, case-insensitively;
- `configuration.path`;
- `configuration.repository.id`;
- `configuration.repository.defaultBranch`.

For an exact-name match, this comparison is performed against the complete
Build Definition returned by `GET /_apis/build/definitions/{id}`. The target
Azure DevOps Server's Pipelines by-ID model can omit
`configuration.repository.defaultBranch`; treating that missing field as a
mismatch would cause a needless Build Definition PUT and revision increment on
every otherwise-idempotent rerun.

When no exact-name Pipeline exists, Build Definitions are filtered by generated
repository ID and `process.yamlFilename`. A legacy same-file match is renamed
instead of creating a duplicate.

The target server rejects `PUT /_apis/pipelines/{id}` with HTTP 405. Therefore
updates follow the official Build Definitions GET-modify-PUT pattern:

1. GET the complete Build Definition.
2. Preserve its ID, revision, and other server fields.
3. Set `name`, `path`, `process.yamlFilename`, repository ID/name/type, and
   default branch.
4. PUT the complete definition to `/_apis/build/definitions/{id}` using API
   `7.1`.

Create and update responses must contain a truthy definition ID; a 2xx response
without an ID is treated as a failure.

The Pipeline API resolves the referenced file from the configured repository
and branch. The YAML must exist on `refs/heads/main` before Pipeline POST or
Build Definition PUT. A
Pipeline folder such as `\komodo` is an Azure DevOps UI/security folder and has
no relationship to the repository path.

## Classic Release definition contract

The generated definition has:

- path from `releaseConfig.folder`, default `\komodo`;
- one primary Build artifact with `sourceId = <projectId>:<pipelineId>`;
- artifact default branch `refs/heads/main` and default version type Latest;
- one environment from `releaseConfig.environmentName`;
- one agent-based deployment phase using the selected queue;
- one Bash task with task GUID
  `6c731c3c-3c68-459a-a5c9-bde6e6595b5b`, version `3.*`, inline target;
- top-level `variableGroups: [<KomodoAPI-id>]`, while the environment-level
  `variableGroups` array stays empty;
- no Release triggers;
- a `ReleaseStarted` environment event condition;
- automated pre- and post-deployment approval records;
- retention of 30 days and three releases, retaining the Build artifact.

The approval objects are required by the target Azure DevOps Server. Empty
approval arrays or null approval objects fail with `VS402877`. The relevant
invariants are:

```json
{
  "preDeployApprovals": {
    "approvals": [
      { "rank": 1, "isAutomated": true, "isNotificationOn": false }
    ],
    "approvalOptions": { "executionOrder": "beforeGates" }
  },
  "postDeployApprovals": {
    "approvals": [
      { "rank": 1, "isAutomated": true, "isNotificationOn": false }
    ],
    "approvalOptions": { "executionOrder": "afterSuccessfulGates" }
  }
}
```

The actual payload includes the remaining approval options required by the
server. Keep the browser and shell payload builders synchronized when changing
this contract.

Before creating or updating a Release, the browser queries `KomodoAPI` by exact
name with `actionFilter=Use`. It validates the presence of `AZP_TOKEN`,
`KOMODO_API_KEY`, and `KOMODO_API_SECRET`, but never logs or persists their
values. Azure DevOps Release definitions link Variable Groups by numeric ID,
not by name. On update, the generator adds the desired ID while retaining any
other existing definition-level group IDs.

The default `packagedFile` script source is loaded from the installed
addressable asset `dist/release-inline-task.sh`, then copied verbatim into the
Bash task's `inputs.script`; the task remains `targetType: inline`. Azure DevOps
variable macros are expanded only when the Release job runs. The wrapper must
not enable shell xtrace around a Git or HTTP Authorization header or place that
header in a process argument.

The browser reconciles Releases rather than treating any same-name definition
as final. It first searches the desired exact name
`<exact-generated-yaml-filename>.yml_Release`; if absent, it inspects expanded
Build artifacts filtered by `artifactType=Build` and
`artifactSourceId=<projectId>:<pipelineId>` for a legacy definition that
references the same Pipeline ID.
The full definition is compared and, when needed, updated with its current ID,
revision, and environment ID preserved. Duplicate-name conflicts are re-read
and reconciled. Historical unrelated definitions are never deleted
automatically.

## API versions: browser versus shell

The browser uses fixed versions close to each call:

- Git, queues, and service endpoints: `6.0`.
- Variable Groups: `7.1`.
- Pipeline create/read: `7.1-preview.1`.
- Build Definition migration/update: `7.1`.
- classic Release: `7.1-preview.4`.

The shell uses `API_VERSION` for projects, repositories, queues, and Pipelines;
its default is `7.1`. It uses `RELEASE_API_VERSION`, default
`7.1-preview.4`, for Release operations. Pipeline creation still appends the
required `repositoryId` query parameter.

This difference is intentional in the current code. Do not document the shell
as using `7.1-preview.1` unless its default is changed. When targeting another
Azure DevOps Server version, verify the exact supported contract with a live
non-production project before changing either path.

## Error contract

Every browser request checks `response.ok`. Non-2xx responses are converted to
errors containing:

- operation description;
- HTTP status;
- sanitized and truncated response body;
- response URL;
- Pipeline or Release domain when relevant.

The UI treats status 401, 403, TF400813, and matching 401 messages as
authorization failures. It adds a domain-specific permission hint but does not
retry REST writes automatically.

HTTP 404 has special meaning only in two Git reads:

- a missing branch ref becomes the zero object ID;
- a missing YAML item means the push should use `add`.

Other 404 responses are failures.

REST list responses are consumed from `value` only. Neither the browser nor the
shell follows continuation tokens, and name lookup returns the first exact
match. Projects with large result sets or duplicate names in different folders
need pagination/disambiguation before this behavior can be considered robust.

## On-premises operational requirements

### Internal proxy bypass

In the verified environment, system HTTP proxy variables route internal Azure
DevOps requests to a proxy that returns nginx 403. Terminal tools must bypass
the proxy for the Azure DevOps hostname/IP, for example with `NO_PROXY`. The
browser normally follows workstation/network policy and is not controlled by
the shell variables.

### TLS trust

Install and trust the internal CA used by Azure DevOps. Do not make `curl -k` or
global TLS verification disabling part of normal automation. Temporary insecure
diagnostics can hide certificate and interception problems.

### File visibility before Pipeline creation

Do not create the Pipeline until the Git push succeeds and the exact YAML path
is readable on the configured branch. `YamlFileNotFoundException` identifies a
repository/path/branch/commit mismatch, not a Pipeline folder problem.

### Extension installation state

An extension can remain installed with hidden `error` or
`needsReauthorization` flags even when a normal by-name GET returns 404. Query
the installed-extension list with disabled/error entries included before
assuming the extension is absent. Recovery details and the verified endpoint
sequence are in the E2E context.

## Readback verification

A successful create response is necessary but not sufficient for a live E2E
test. Read back:

1. the YAML item on `main`;
2. the Pipeline through the Pipelines API;
3. the corresponding classic Build Definition and `process.yamlFilename`;
4. the Release definition, artifact Pipeline ID, repository ID, environment,
   queue, Bash task, and automated approvals.

Do not queue a run merely to verify definition linkage unless execution was
explicitly requested.

## Cleanup endpoints

On the verified Azure DevOps Server, deleting through
`DELETE .../_apis/pipelines/{id}` returned HTTP 405. Use the Build Definitions
endpoint for a disposable YAML Pipeline:

```text
DELETE <project>/_apis/build/definitions/{id}?api-version=7.1
```

Delete a disposable classic Release definition with:

```text
DELETE <project>/_apis/release/definitions/{id}?api-version=7.1-preview.4
```

Resolve exact IDs first and do not delete the preserved resources recorded in
`azure-devops-e2e-state.yaml` without explicit approval.

## Official references

These Microsoft references describe the public Azure DevOps Services contracts.
The on-premises compatibility requirements documented above are additional
observations from the target Azure DevOps Server and may intentionally differ
in API version or URL shape.

- [Authenticate and secure Azure DevOps web extensions](https://learn.microsoft.com/en-us/azure/devops/extend/develop/auth?view=azure-devops)
- [Navigate the parent Azure DevOps host from an extension](https://learn.microsoft.com/en-us/azure/devops/extend/develop/work-with-urls?view=azure-devops)
- [Sign in to Azure DevOps Server with different credentials](https://learn.microsoft.com/en-us/azure/devops/organizations/projects/connect-to-projects?view=azure-devops)
- [Azure DevOps extension manifest and scopes](https://learn.microsoft.com/en-us/azure/devops/extend/develop/manifest?view=azure-devops)
- [Create a Pipeline](https://learn.microsoft.com/en-us/rest/api/azure/devops/pipelines/pipelines/create?view=azure-devops-rest-7.1)
- [Update a Build Definition with its current revision](https://learn.microsoft.com/en-us/rest/api/azure/devops/build/definitions/update?view=azure-devops-rest-7.1)
- [Create a Git push](https://learn.microsoft.com/en-us/rest/api/azure/devops/git/pushes/create?view=azure-devops-rest-7.1)
- [Create a classic Release definition](https://learn.microsoft.com/en-us/rest/api/azure/devops/release/definitions/create?view=azure-devops-rest-7.1)
- [Update a classic Release definition](https://learn.microsoft.com/en-us/rest/api/azure/devops/release/definitions/update?view=azure-devops-rest-7.1)
- [List classic Release definitions and continuation-token parameters](https://learn.microsoft.com/en-us/rest/api/azure/devops/release/definitions/list?view=azure-devops-rest-7.1)
- [Get Variable Groups by project and name](https://learn.microsoft.com/en-us/rest/api/azure/devops/distributedtask/variablegroups/get-variable-groups?view=azure-devops-rest-7.1)
