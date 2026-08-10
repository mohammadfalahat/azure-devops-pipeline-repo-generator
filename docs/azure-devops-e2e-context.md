# Azure DevOps end-to-end context

Last complete five-step provisioning verification: **2026-08-10**

Complete five-step workflow verified version: **0.1.30**

Latest installed launch behavior observed: **0.1.30 on 2026-08-10**

Local Release-script/Variable-Group candidate awaiting live verification: **0.1.32**

This document records the durable findings from debugging and testing the
Pipeline Generator extension against the on-premises Azure DevOps instance. It
is deliberately layered: start at Layer 0 and read deeper only when the current
task needs it. The companion
[`azure-devops-e2e-state.yaml`](azure-devops-e2e-state.yaml) is the canonical
machine-readable snapshot.

No credentials belong in either file. A PAT was exposed during the original
conversation and must never be copied from chat into commands, documentation,
or logs; rotate any exposed PAT.

## Current local candidate — version 0.1.32

Version 0.1.32 carries forward the locally tested no-write fixes from 0.1.31
and changes the desired classic Release definition in two deliberate ways:

- packaged asset `dist/release-inline-task.sh` contains the requested Komodo
  bootstrap wrapper; the browser reads this asset when provisioning and stores
  its full content as the Bash@3 task's Inline script;
- the browser resolves exact project Variable Group `KomodoAPI`, verifies the
  presence of `AZP_TOKEN`, `KOMODO_API_KEY`, and `KOMODO_API_SECRET`, and links
  its numeric ID at Release-definition scope while preserving any additional
  existing definition-level groups.

The wrapper corrects the supplied ERR trap's extra closing brace and never
enables shell xtrace around the secret-derived Basic Authorization header or
places it in a process argument. No credential value is present in the packaged
asset; Azure DevOps variable macros are expanded only when the Release job runs.

A read-only check from the normal signed-in parent Chrome session on 2026-08-10
used no PAT or authentication adaptation and established the live server shape:

- `GET .../variablegroups?groupName=KomodoAPI&actionFilter=Use` returned HTTP
  200, exact group ID 7, type `Vsts`, and the three required variable names;
  all three were marked secret and no values were inspected or logged;
- working reference Release definition 3 (`New UI Pro`) returned HTTP 200 with
  top-level `variableGroups: [7]` and environment-level `variableGroups: []`.

This is contract evidence, not a write test of candidate 0.1.32. The manifest
now requests `vso.variablegroups_read`, so the collection installation must be
updated/reauthorized before the first hosted run. That first run is expected to
update preserved Release definition 7 once because both its Inline script and
Variable Group linkage change. It must not alter the byte-identical YAML or
Pipeline 347 revision 2. A second 0.1.32 run should then perform no Release PUT.

Candidate artifact:

- File: `mohammad-falahat.pipeline-generator-0.1.32.vsix`
- SHA-256:
  `28419f876deefc904ca0bd8b8f234cc6e8297d3580e0d1a1c3e7e6215328a3c3`.
- Full local suite, Bash syntax, structured E2E-state YAML parse, packaged
  source equality, VSIX manifest/version/scope inspection, and package-source
  comparisons: passed.
- Installed/live Release reconciliation status: **pending**.

## Live completion — installed 0.1.30; no-write fix in candidate 0.1.31

A normal signed-in Chrome run of installed 0.1.30 from `feature/Zones` completed
all five steps. This was not a PAT-adapted test: the generator obtained its
host token normally, and the result was monitored through the browser's
loopback DevTools endpoint. Independent read-only REST readback used the same
signed-in parent session.

The scope correction is verified. The installed-extension API reported version
0.1.30, `installState.flags=none`, and both `vso.agentpools` and
`vso.serviceendpoint`. Step 5 resolved queue `PublishDockerAgent` as ID 111 and
created classic Release definition 7 named
`ridesharing-ridesharing_backend-feature-zones.yml_Release` in `\komodo`.
Its Build artifact references Pipeline 347, its environment is `komodo`, its
Bash 3 task contains the configured inline script, its `ReleaseStarted`
condition and automated pre/post approvals are present, and its revision is 1.
No Release instance was created.

The resource readback also proved:

- the YAML blob content was unchanged and its last path-specific content change
  remains commit `4cde26160ce6b2b222f0925e8c53b35ed373ef34` from 2026-08-09;
- exactly one matching Pipeline exists: ID 347, exact filename name, folder
  `\KOMODO`, generated repository ID, `refs/heads/main`, and the exact YAML path;
- no Build was queued for Pipeline 347 and no Release instance exists for
  definition 7;
- preserved Pipeline 344 and Release 5 were not deleted.

Two stricter idempotency defects were exposed. Step 2 sent an `edit` Push for
the byte-identical YAML, advancing `main` with empty commit
`726f11e85fd78adecc5aefaece8686ecb0282aa2`; its changes endpoint returns no
changed paths. Pipeline 347 was reused rather than duplicated, but its Build
Definition revision also increased from 1 to 2. The target Server omits
`repository.defaultBranch` from its sparse Pipelines by-ID response, and
version 0.1.30 treated that omission as a binding mismatch.

Candidate 0.1.31 reads existing YAML text and skips the Git Push when content
is equal. It also reads the canonical full Build Definition for an exact-name
Pipeline before comparing the binding. A correct rerun therefore returns
without a YAML commit, Build Definition PUT, or revision increment. Regression
tests model both no-write paths.

Superseded local 0.1.31 artifact (never installed):

- File: `mohammad-falahat.pipeline-generator-0.1.31.vsix`
- SHA-256:
  `ddc364bd2c3392719b0b0ea9634d71d8817e14abbd5d6ba9fff0b8528e6dfc9f`.
- Full local suite, unchanged-YAML no-push and strict Pipeline no-op
  regressions, structured YAML parsing, VSIX integrity, packaged-source
  comparison, and packaged-scope inspection: passed.
- Installed/live no-write rerun status: **pending**. The acceptance baseline is
  Pipeline 347 revision 2 and Release 7 revision 1; both must remain unchanged.

## Previous scope correction — version 0.1.30

Installed 0.1.28 had reached Step 5 but its extension token received HTTP 401
`TF400813` while loading queue `PublishDockerAgent`; the same user's parent
session could read queues and Docker Registry endpoints with HTTP 200. Version
0.1.30 added the least-privilege read scopes `vso.agentpools` and
`vso.serviceendpoint`, tagged missing-scope failures, and directed collection
administrators to extension authorization. The 2026-08-10 normal signed-in run
above verifies that this root-cause correction works.

## Previous local candidate — version 0.1.29

The installed 0.1.28 screenshot proves the Hub width remains stable and the
submit button can be rendered, but the user reports that mouse-wheel scrolling
does nothing. Version 0.1.28 assigned `overflow-y: auto` to `body` while the
root `html` element used hidden overflow. The target legacy iframe suppresses
that root/body scroll combination.

Version 0.1.29 keeps both root elements bounded and moves scrolling to an
explicit `.wrapper` container with `position: fixed`, `inset: 0`, `height:
100%`, and `overflow-y: scroll`. The scrollbar is owned entirely by the
extension viewport and does not require `VSS.resize()` or host-page scrolling.

Candidate artifact:

- File: `mohammad-falahat.pipeline-generator-0.1.29.vsix`
- SHA-256:
  `4a3fb63791c1996e7acdda546f079d51ed26906438e178684d10f22d558a4366`.
- Local regression suite, structured YAML parsing, 1600×600 visual overflow
  check, VSIX integrity, and packaged-source comparisons: passed.
- Installed/live scrolling status: **not yet verified**.

## Installed partial verification — version 0.1.28

The installed 0.1.27 screenshot shows a severe horizontal feedback loop: after
the Hub opened, blank space on the right kept expanding while the generator
iframe collapsed into a narrow strip on the left. The trigger was 0.1.27's
parameterless `VSS.resize()` call. In the bundled legacy SDK, omitted dimensions
are replaced with `body.scrollWidth` and `body.scrollHeight`; repeatedly feeding
the changing `scrollWidth` back to this Azure DevOps Server caused the host and
iframe widths to amplify each other.

Version 0.1.28 removes all host resize calls and resize observers. The Hub
document is instead bounded to the host viewport: `html` owns a fixed 100%
height with hidden outer overflow, while `body` owns stable internal vertical
scrolling. The wrapper uses `box-sizing: border-box` and `width: 100%`, so the
complete form remains reachable without changing contribution width.

The authorization recovery introduced in 0.1.27 remains. Its message handling
is also corrected so normalized `HostAuthorizationNotFound` errors always lead
with **Open extension authorization** and explicitly explain that signing out
cannot create a missing collection-level extension authorization record.

Candidate artifact:

- File: `mohammad-falahat.pipeline-generator-0.1.28.vsix`
- SHA-256:
  `9d474aa224658112149a2e24fc36a137092eb670b580f068bb43660cabefdb4b`.
- Local extension, action, UI, listener, shell self-tests, structured YAML
  parsing, package integrity, packaged-runtime comparison, and the no-resize
  package assertion: passed.
- Installed Hub width/layout: **verified interactively as stable**.
- Complete vertical reach to the submit button: **not yet verified**.
- After the user's clean reinstall, direct token issuance: **succeeded**.
- Provisioning reached Step 5, then failed while reading the Release queue
  because the 0.1.28 manifest omitted `vso.agentpools`.

The clean reinstall resolved the missing host-authorization record. A later
Chrome DevTools inspection confirmed that the hosted page obtained a token and
created Pipeline 347, but its token could not read agent queues or service
endpoints because those resource-area scopes were not declared in 0.1.28.

## Observed installed launch — version 0.1.27

Version 0.1.27 was installed successfully and the collection extension details
page displayed all requested Code, Project, Build, and Release permissions. The
Hub still opened in-host with correct branch context, but the new automatic
resize behavior caused its width to shrink continuously while a blank host
region grew from the right. The form was eventually compressed to an unusable
narrow column.

Installed artifact:

- File: `mohammad-falahat.pipeline-generator-0.1.27.vsix`
- SHA-256:
  `5829e4edaae95800844333158f12d6d6033f086b5c3feb6403630cb7217b0353`
- Extension update and requested permissions display: **verified
  interactively**.
- Hub layout: **failed because of a `VSS.resize()` width feedback loop**.
- Direct host token issuance and provisioning: **not verified in this run**.

## Observed installed launch — version 0.1.26

Version 0.1.26 removed detached launch completely and declared
`pipeline-generator-hub` as an `ms.vss-web.hub` targeting
`ms.vss-code-web.code-hub-group`. When `openCustomDialog` was unavailable, the
action used the host navigation service to open that Hub in the current Azure
DevOps page.

The user's screenshot confirms that this in-host fallback succeeded and that
the branch context was transferred correctly. It also records the two failures
that motivated 0.1.27: the bottom of the form was clipped, and the hosted page
received `HostAuthorizationNotFound` from `VSS.getAccessToken()`.

Installed artifact:

- File: `mohammad-falahat.pipeline-generator-0.1.26.vsix`
- SHA-256:
  `5088d833572ec6aa6c8ae2fafbd81e3dbf71aeccda032791693eb9a48b3019b4`
- In-host Hub launch and branch context: **verified interactively**.
- Complete form visibility: **failed**.
- Direct host token issuance: **failed with `HostAuthorizationNotFound`**.
- Complete provisioning: **not attempted because authorization was absent**.

This observation proves that an in-host contribution alone does not resolve a
missing extension authorization record.

## Previous development delta — candidate 0.1.26

The user's interactive screenshots prove that Azure DevOps served the 0.1.25
asset, but selecting **Generate pipeline** still opened this URL in a separate
browser tab:

```text
/_apis/public/gallery/publisher/mohammad-falahat/extension/
pipeline-generator/0.1.25/assetbyname/dist/index.html?branch=...
```

That asset URL plus query string is the exact signature of the 0.1.25 detached
fallback. Therefore the custom Dialog attempt did not succeed on this Azure
DevOps Server. The screenshot does not expose the caught service exception, so
the durable conclusion is limited to **Dialog unavailable/failed and fallback
executed**; do not claim a more specific server error without console evidence.

Version 0.1.26 removes detached launch completely. It also declares
`pipeline-generator-hub` as an `ms.vss-web.hub` targeting
`ms.vss-code-web.code-hub-group`. The action still prefers
`openCustomDialog`; if unavailable, it uses the legacy host navigation service
to navigate the current Azure DevOps page to:

```text
<collection>/<project>/_apps/hub/
mohammad-falahat.pipeline-generator.pipeline-generator-hub?branch=...
```

The Hub hosts `dist/index.html` inside Azure DevOps. The UI reads branch and
repository query values from `getCurrentState()` (legacy service) or
`getQueryParams()` (modern service), then calls `VSS.getAccessToken()` from its
own host iframe. The branch action has no token-acquisition, `window.open`,
`openWindow`, or `postMessage` bootstrap path.

Candidate artifact:

- File: `mohammad-falahat.pipeline-generator-0.1.26.vsix`
- SHA-256:
  `5088d833572ec6aa6c8ae2fafbd81e3dbf71aeccda032791693eb9a48b3019b4`
- Local syntax, Dialog/Hub behavior, package-content checks, and `npm test`:
  passed.
- Installed/live E2E status: subsequently observed as described above.

After installation, either a modal or the Pipeline Generator page under Azure
Repos is valid. A separate tab is not valid. The installed result selected the
Repos Hub path; direct token acquisition then failed because extension
authorization was unavailable.

## Observed installed launch — version 0.1.25

Version 0.1.25 changes the generator launch architecture in response to the
interactive finding that the action and detached generator could report
`HostAuthorizationNotFound` even while the user's visible Azure DevOps session
was still active.

The manifest now declares `pipeline-generator-dialog` as an
`ms.vss-web.control` whose content is `dist/index.html`. The branch action asks
`ms.vss-features.host-page-layout-service` to open the fully-qualified control
ID with `openCustomDialog`. Branch, project, repository, and collection values
are passed as non-secret dialog configuration. Inside that Azure DevOps-hosted
iframe, `ui.js` reads `VSS.getConfiguration()` and calls
`VSS.getAccessToken()` itself. The normal path therefore neither opens a new
tab nor transfers a token between extension pages.

The old detached-window launch remains only as a compatibility fallback for a
host without `openCustomDialog`; only that fallback requests a token in the
action and transfers it through same-origin `postMessage`. After successful
provisioning the generator now uses host navigation for the Pipeline URL, so
the parent Azure DevOps page moves to the Pipeline instead of navigating only
the dialog iframe.

Installed artifact originally supplied:

- File: `mohammad-falahat.pipeline-generator-0.1.25.vsix`
- SHA-256:
  `9abafb13174ebe55d3f5ebd5e5334c86c1d310e8f9552d74a181d374a07d265c`
- Local syntax, behavior, package-content checks, and `npm test`: passed.
- Installed launch status: **fallback to a standalone tab observed**.
- Complete provisioning and direct in-frame token issuance: **not verified**.

This result proves that treating a detached tab as an acceptable compatibility
fallback was incorrect for the target Server: it cannot perform the generator
host handshake, and the action-side token request still returned
`HostAuthorizationNotFound`.

## Previous development delta — candidate 0.1.24

Version 0.1.24 includes the Pipeline/Release reconciliation fixes prepared in
0.1.22 and adds the requested recovery path for the newly reported
authentication failure.
In the user's interactive browser, both the action and generator received
`HostAuthorizationNotFound`; before and after clicking Create, version 0.1.22
only displayed an access-token error and could not proceed.

An intermediate 0.1.23 candidate added a page-memory PAT fallback, but it was
superseded before live verification because the required recovery is to destroy
the current Azure DevOps browser session and run the full login flow again.
Version 0.1.24 never asks for a PAT. When token acquisition fails, it shows an
explicit **Sign out and authenticate again** action, clears the in-memory host
token, and navigates the parent host to the collection-relative `_signout`
route using the Azure DevOps navigation service (with top-level navigation as
fallback). After login, the user reopens Generate pipeline from the branch.

Read-only anonymous probes on 2026-08-09 confirmed that both deployment-root
and collection-relative `_signout` routes are recognized by the target server
and respond with its normal authentication challenge (`Bearer`, `Basic`,
`Negotiate`, and `NTLM`). They do not prove cookie invalidation in the user's
authenticated browser; that remains an installed interactive test.

The earlier 0.1.22 changes were created after a retry exposed two defects in
0.1.21:

- successful Pipeline responses could still be followed by unsupported
  `PUT /_apis/pipelines/{id}` when the server normalized folder casing;
- after a failed/partial submit, generated repository metadata overwrote source
  repository state, changing the next Pipeline name to include
  `RideSharing_Azure_DevOps`.

The candidate keeps source and generated repository identities separate and
sets the Pipeline display name to the exact generated YAML filename. It finds a
legacy Pipeline by repository/YAML path and performs the supported Build
Definition GET-modify-PUT with the current revision. Release lookup now uses
both exact desired name and Pipeline artifact ID; matching legacy definitions
are reconciled instead of duplicated.

Candidate artifact:

- File: `mohammad-falahat.pipeline-generator-0.1.24.vsix`
- SHA-256:
  `8ac776bc11b9504447166892e0bb4df72fef1495d1a979bdbf7e85fd2c4b0290`
- Local syntax, package-content checks, and `npm test`: passed.
- Installed/live E2E status: **not yet verified**.

The older 0.1.21 resource graph remains useful historical evidence and is kept
alongside the newer feature/Zones graph. Do not attribute either live result to
an uninstalled local candidate.

## Layer 0 — Current answer and quick facts

The installed **Pipeline Generator 0.1.30** completed its actual hosted UI flow
for `feature/Zones` and left all three linked resource types present:

1. YAML file in `RideSharing_Azure_DevOps` on `main`.
2. Real YAML pipeline linked to that exact file.
3. Classic Release definition whose build artifact is that pipeline.

The generated resources are:

| Resource | Verified value |
| --- | --- |
| YAML | `/ridesharing-ridesharing_backend-feature-zones.yml`, object ID `1c5181a428f7f7a536ddbd331e1e1230a6a23f77`; content unchanged, but 0.1.30 created an empty no-op commit |
| Pipeline | ID `347`, `ridesharing-ridesharing_backend-feature-zones.yml`, folder `\KOMODO`, revision `2` |
| Release definition | ID `7`, `ridesharing-ridesharing_backend-feature-zones.yml_Release`, folder `\komodo`, revision `1` |
| Execution side effects | Build count `0`; Release-instance count `0` |

These are the latest verification resources. **Do not delete Pipeline IDs 344
or 347, or Release-definition IDs 5 or 7, as routine cleanup.** On the first
0.1.32 run, Pipeline 347 must stay at revision 2 while Release 7 is expected to
advance once from revision 1 to receive the new Inline wrapper and `KomodoAPI`
linkage. A second run must leave both revisions unchanged.

Useful links:

- [Generated feature/Zones YAML](https://azure.buluttakin.com/ShonizCollection/RideSharing/_git/RideSharing_Azure_DevOps?path=%2Fridesharing-ridesharing_backend-feature-zones.yml&version=GBmain)
- [Pipeline 347](https://azure.buluttakin.com/ShonizCollection/RideSharing/_build?definitionId=347)
- [Release definition 7](https://azure.buluttakin.com/ShonizCollection/RideSharing/_release?definitionId=7)

## Layer 1 — Environment and verified object graph

### Azure DevOps topology

- Server: `https://azure.buluttakin.com`
- Internal address observed: `192.168.62.17`
- Collection: `ShonizCollection`
- Project: `RideSharing`
- Project ID: `0eaa21d3-9359-45a7-a374-6f44bde2d941`
- Source repository: `RideSharing_Backend`
- Source repository ID: `72f0181d-3763-4df4-b995-1a4701791c3b`
- Latest tested source branch: `feature/Zones`
- Older retained verification branch: `feature/defineZones`
- Generated repository: `RideSharing_Azure_DevOps`
- Generated repository ID: `316fba81-11a5-4dfe-a71f-aa91486e4279`
- Generated repository default branch: `refs/heads/main`
- Release agent queue: `PublishDockerAgent`, queue ID `111`

### Verified linkage

```text
RideSharing_Backend @ feature/Zones
  -> RideSharing_Azure_DevOps @ refs/heads/main
     -> /ridesharing-ridesharing_backend-feature-zones.yml
        -> YAML pipeline 347
           -> Classic Release definition 7
              -> environment komodo / queue 111 / inline Bash task
```

Pipeline 347 was verified through both the Pipelines response and the classic
Build Definition response:

- `configuration.type` is `yaml`.
- `configuration.path` is the exact generated YAML path.
- `configuration.repository.id` is the generated repository ID.
- `repository.defaultBranch` is `refs/heads/main`.
- Build Definition `process.yamlFilename` is the exact generated YAML path.
- The server normalizes the requested pipeline folder `\komodo` to `\KOMODO`.

Release definition 7 was verified to contain:

- Pipeline artifact definition ID `"347"` and the exact filename-based Pipeline name.
- Target generated repository ID `316fba81-11a5-4dfe-a71f-aa91486e4279`.
- Environment `komodo` using queue ID `111`.
- Bash task ID `6c731c3c-3c68-459a-a5c9-bde6e6595b5b`, version `3.*`, inline target.
- Automated pre-deployment and post-deployment approval entries.
- No Release instance.

That 0.1.30 definition predates the 0.1.32 wrapper/Variable-Group contract; do
not treat its current revision 1 as already carrying `KomodoAPI`.

The older Pipeline 344 and Release definition 5 from the 0.1.21
`feature/defineZones` verification remain preserved historical resources.

An older file, `/ridesharing-ridesharing_backend-demo.yml` with object ID
`8ab4e461ad71872ba129af1029765172fd605557`, remains from version 0.1.16 and is
not evidence for the corrected branch-specific flow.

## Layer 2 — Root causes and implemented contracts

### Pipeline creation on Azure DevOps Server

The installed browser UI uses the Azure DevOps Server-compatible Pipelines API
contract `7.1-preview.1`. Pipeline creation also requires
`repositoryId=<generated-repository-guid>` in the request query string, even
though the same repository ID is already present in the JSON body.

The UI implements this through `PIPELINE_API_VERSION`,
`buildPipelinesApiUrl`, and `repositoryId: repo.id` in `dist/ui.js`. The shell
provisioner also appends the repository query parameter through
`pipeline_collection_url`, but its configurable `API_VERSION` currently
defaults to stable `7.1`. That shell version reached server-side YAML
resolution during diagnostics. Therefore do not claim both execution paths use
the preview version; the required shared invariant is the repository binding in
the create URL.

Pipeline create/update responses must contain an ID. The UI validates this via
`readPipelineResponse`; a successful HTTP status without a pipeline ID is not
accepted as success.

### Classic Release approvals

The first Release POST failed with:

```text
VS402877: Pre-approvals or post-approvals in stage 'komodo' are empty.
```

This Azure DevOps Server rejects empty approval arrays or null approvals. Each
side must have an automated entry:

```json
{
  "approvals": [
    { "rank": 1, "isAutomated": true, "isNotificationOn": false }
  ]
}
```

Pre-approval options use `executionOrder: "beforeGates"`; post-approval options
use `executionOrder: "afterSuccessfulGates"`. The environment conditions also
include:

```json
{
  "name": "ReleaseStarted",
  "conditionType": "event",
  "value": "",
  "result": null
}
```

Both `dist/ui.js` and `scripts/provision-pipeline-release.sh` implement these
requirements. `scripts/validate-extension.js` asserts the pipeline URL binding
and approval shape. The local `npm test` suite passed after the changes.

### Extension package and installation recovery

- Publisher: `mohammad-falahat`
- Extension ID: `pipeline-generator`
- Display name: `Pipeline Generator`
- Verified installed version: `0.1.21`
- Package: `mohammad-falahat.pipeline-generator-0.1.21.vsix`
- Package SHA-256:
  `44111fc1e7f2184df4b7406dfb8dddc94b6f70b50eee34ead600c3ec09ce5e4f`

A stale 0.1.16 installation was hidden in state flags
`error, needsReauthorization`. The ordinary installed-extension GET returned
404, while the full list with `includeDisabledExtensions=true` and
`includeErrors=true` revealed it. Installing 0.1.21 directly then returned 409
because Azure DevOps still considered the stale extension installed.

The successful recovery was to delete the stale collection installation and
then clean-install version 0.1.21:

```text
DELETE /ShonizCollection/_apis/extensionmanagement/installedextensionsbyname/
       mohammad-falahat/pipeline-generator?api-version=7.1-preview.1

POST   /ShonizCollection/_apis/extensionmanagement/installedextensionsbyname/
       mohammad-falahat/pipeline-generator/0.1.21?api-version=7.1-preview.1
```

The POST has an empty body and `Content-Length: 0`. After reinstall,
`installState.flags` was `none`.

## Layer 3 — What the UI test proves

The test started from the real Branches page, selected the More options menu on
`feature/defineZones`, opened **Generate pipeline**, loaded the installed
0.1.21 asset, and exercised the page's real `dist/ui.js` logic. The UI reached:

```text
Done. Pipeline RideSharing_RideSharing_Backend_demo is linked to
/ridesharing-ridesharing_backend-feature-definezones.yml in \komodo (ID: 344).
Release definition created (ID: 5). Opening Pipelines...
```

The redirect went to the actual pipeline page, and independent REST reads
verified the object graph in Layer 1.

There is one important authentication boundary. A headless Chrome page loaded
with PAT Basic authentication cannot obtain the short-lived extension token
from `VSS.getAccessToken()`; the host returns `Error issuing session token:
AccessDenied`. To test the installed UI logic, the test injected the same
`pipeline-bootstrap` branch payload and adapted its requests from
`Authorization: Bearer <PAT>` to Basic authentication. Therefore this is an
**authentication-adapted installed-UI test**, plus independent API
verification—not a pure automated proof of token issuance inside the user's
already signed-in interactive Chrome session.

This boundary does not invalidate the verified create/link workflow. If a
future task specifically requires an exact interactive-session proof, have the
user click in their already authenticated browser while taking REST snapshots
before and after, or launch an automation-enabled browser with a real
interactive Azure DevOps login from the beginning.

## Layer 4 — Safe operational procedures

### Internal-host proxy bypass

The execution environment defines HTTP proxy variables. Sending this internal
Azure DevOps traffic through that proxy produced nginx 403 responses. Use
`curl --noproxy '*'` or run the test process with the proxy variables unset and
`NO_PROXY=azure.buluttakin.com,192.168.62.17`.

### Temporary PAT handoff

Never put the PAT directly on a command line. If live verification is
authorized, ask the user to create a mode-0600 temporary file without echoing
the token:

```bash
umask 077
read -rsp 'Azure DevOps PAT: ' CODEX_UI_PAT
printf '%s' "$CODEX_UI_PAT" > "/tmp/codex-azp-token-$(id -u)"
unset CODEX_UI_PAT
echo
```

Read the token only inside the process that needs it, do not print it, and
remove the file with `shred -u` when finished. Do not embed a PAT in a URL,
shell history, test artifact, browser screenshot, or repository file.

The tested PAT could manage the relevant project APIs and extension
installation. It received 401 from `tfx extension show/publish` against the
Marketplace, so do not assume that the same PAT can publish packages.

### Correct deletion endpoints for disposable tests

On this server, `DELETE /_apis/pipelines/{id}` returned HTTP 405. Delete a
disposable YAML pipeline through the classic Build Definitions endpoint:

```text
DELETE /ShonizCollection/RideSharing/_apis/build/definitions/{id}?api-version=7.1
```

Delete a disposable Release definition with:

```text
DELETE /ShonizCollection/RideSharing/_apis/release/definitions/{id}?api-version=7.1-preview.4
```

Earlier disposable pipeline 343 and release definition 4 were deleted this
way. Do not apply cleanup to the current verified pipeline 344 and release 5
without explicit user approval.

## Layer 5 — Known dead ends; do not repeat by default

- No native Playwright/browser-control tool was available.
- The already-running Chrome exposed no DevTools remote-debugging endpoint.
- Copying Chrome cookies into a temporary headless profile produced
  `ERR_INVALID_AUTH_CREDENTIALS`; the on-premises login is not reproducible from
  cookies alone.
- PAT Basic authentication loads the page but cannot obtain the VSS extension
  session token.
- Wayland prevented `xdotool`/`wmctrl`-style focus and input control.
- GNOME Shell D-Bus calls for window discovery, screenshots, focus, or a remote
  desktop session were denied or inhibited.
- AT-SPI could enumerate the Chrome application and frame, but could not focus
  the pre-existing Wayland window; its renderer accessibility tree was not
  available because Chrome had started before accessibility was enabled.
- Toolkit accessibility was temporarily enabled during investigation and was
  restored to `false`. Temporary browser profiles, helper files, and token
  files were removed.

The non-sensitive success screenshot was written to
`/tmp/pipeline-generator-0.1.21-ui-success.png`. `/tmp` is ephemeral; it is not
durable evidence and should not be referenced as if it were versioned.

## Layer 6 — Area-project diagnostic distinction

An earlier shell command targeted project `Area`:

- Project ID: `82f378a4-46ca-4fa8-88ab-6cc2c9d4c7fd`
- Generated repository: `Area_Azure_DevOps`
- Generated repository ID: `b5255b04-a863-4712-a6fe-ec548a6cc468`

At diagnostic time, its `main` branch contained
`/area-area.api-main.yml`, but not `/area-area-definezones.yml`. Pipeline
creation therefore correctly failed with `YamlFileNotFoundException` at commit
`d4928291edca8e204f1b344ab92a3f90bccf2a10`. The pipeline folder `\komodo`
does not affect or substitute for the repository YAML path. Do not conflate
this missing-file failure with the corrected RideSharing end-to-end result.

## Layer 7 — Next-run checklist

1. Read the YAML snapshot and query current resources before creating anything.
2. Confirm the installed extension version and `installState.flags`.
3. Confirm the generated YAML exists on `refs/heads/main` before pipeline POST.
4. For the browser path, confirm Pipeline API `7.1-preview.1`; for both browser
   and shell creation, confirm the URL includes `repositoryId`.
5. Confirm pipeline response has an ID and then read it back through both
   Pipelines and Build Definitions APIs.
6. Confirm Release pre/post approvals are automated and non-empty.
7. Confirm the Release artifact references the created pipeline ID and the
   environment uses queue 111.
8. Record whether authentication was interactive-session, adapted browser, or
   direct REST.
9. For candidate 0.1.32, rerun `feature/Zones` and prove Pipeline 347 stays at
   revision 2; Release definition 7 may advance once from revision 1 only to
   install the new Inline wrapper and `KomodoAPI` ID 7. Run a second time and
   prove the Release revision is then stable. No Build or Release instance may
   appear.
10. If `HostAuthorizationNotFound` appears, use **Open extension
   authorization** as a Collection Administrator, authorize the requested
   scopes, and reopen the generator. If no action is offered, reinstall the
   same published version and recheck `installState.flags`.
11. Record whether the generator opened as an Azure DevOps Dialog or Azure
   Repos Hub and whether its own `VSS.getAccessToken()` succeeded without PAT
   injection or token transfer.
12. Confirm successful provisioning navigates the parent host, not only the
    dialog iframe.
13. Do not queue a run or create a release instance unless explicitly requested.
14. For exact-name Pipeline reuse, confirm the full Build Definition is read
    and no Pipeline or Build Definition PUT is sent when its binding matches.
15. Confirm Release definition top-level `variableGroups` includes ID 7,
   environment-level `variableGroups` remains empty, and the task contains the
   packaged wrapper without logging a secret-derived Authorization header.
16. Update both context files with the date, resulting IDs, and cleanup status.
