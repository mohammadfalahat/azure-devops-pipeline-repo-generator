# Azure DevOps Pipeline Repo Generator

A lightweight Azure DevOps extension that adds a **Generate pipeline** action beside each branch. When invoked, it opens a form pre-populated with the required deployment settings:

- `pool`: `PublishDockerAgent`
- `service`: `api`
- `environment`: `demo`
- `dockerfileDir`: `src/TMS.API`
- `repositoryAddress`: `registry.buluttakin.com`
- `containerRegistryService`: `BulutReg`
- `komodoServer`: `DEMO-192.168.62.91` (with other options available)

The extension uses the short-lived, scoped Azure DevOps access token provided by the host page to create repositories, scaffold the YAML, register the pipeline, and create the classic Release definition—no extra prompts or saved tokens are required. The token represents the signed-in user, so Azure DevOps permission checks are applied to that user; it does **not** bypass repository, pipeline, release, or agent-queue security.

The manifest scopes include `vso.code_manage` so the extension can create repositories on behalf of the signed-in user (who must also have **Create repository** permission in the project).

Submitting the form ensures a shared repository named `<ProjectName>_Azure_DevOps` exists in the current project (using the project name exactly as Azure DevOps reports it). If it does not, the extension creates it and pushes a YAML template named `<ProjectName>-<RepositoryName>-<BranchName>.yml` (preserving the original casing for the project and repository portions) containing the submitted settings.

The branch action targets both the legacy (`git-branches-*`) and the newer repository branches menus to tolerate Azure DevOps UI updates where a single menu surface might go missing.

## Structure

- `vss-extension.json`: Extension manifest defining the branch action and a hub entry.
- `dist/index.html`: Form UI loaded from the action or the hub.
- `dist/ui.js`: Client-side logic that creates the repository/YAML, registers the YAML pipeline, and creates the classic Release definition.
- `dist/release-config.js`: Versioned, token-free configuration for the Release folder/job and predefined inline Bash code (or a Bash file read from Azure Repos).
- `dist/menu-action.js`: Registers the branch menu action and opens the UI with branch context.
- `dist/styles.css`: Lightweight styling for the form.

## Packaging (creating the VSIX)

The extension is static (HTML/CSS/JS only), so packaging is just zipping the manifest and `dist` folder into a VSIX. You can do
this manually or with the official `tfx-cli` utility.

> If you install a new VSIX but do not see the **Generate pipeline** action beside your branches, refresh the Repos page and
> check both the branch context menu and the toolbar menu. The manifest targets multiple branch menus to align with Azure
> DevOps UI variations, so the action should appear in at least one of those locations after the page reloads.

### Prerequisites

- Node.js 18+ (for `npm` and `tfx-cli`).
  - Optional: run `npm install` to pull the Azure DevOps SDK (`vss-web-extension-sdk`) into
    `node_modules/vss-web-extension-sdk/lib`, then copy it to `dist/lib` if you want to bundle the SDK
    with the VSIX for fully offline deployments. If you skip this step, the extension automatically
    loads the SDK from the published extension asset
    (`https://azure.buluttakin.com/_apis/public/gallery/publisher/localdev/extension/pipeline-generator/0.1.11/assetbyname/dist/lib/VSS.SDK.min.js`),
    then falls back to the platform-hosted path `/_content/MS.VSS.SDK/scripts/VSS.SDK.min.js`.
- A publisher ID and display name to embed in `vss-extension.json` (`publisher` and `name` fields). For on-premises servers you
  can use any unique publisher ID (it is not tied to the public marketplace).
- An icon at `dist/images/icon.svg` (a 128x128 SVG) if you want to replace the placeholder.
- The manifest `targets` already include Azure DevOps Services (`Microsoft.VisualStudio.Services.Cloud`) and on-premises
  Azure DevOps Server (`Microsoft.TeamFoundation.Server` with a version range of `[16.0,20.0)` to cover 2019, 2020, and
  2022). If you are targeting an older or newer server release and see a `versionCheckError`, adjust the range accordingly
  before packing.

### Using `tfx-cli`

1. Install the CLI: `npm install -g tfx-cli` (or download the release ZIP if your server blocks npm registry traffic).
2. Bump the `version` in `vss-extension.json` as needed.
3. Run the pack command from the repo root:

   ```bash
   tfx extension create --manifest-globs vss-extension.json --rev-version
   ```

   This outputs a `*.vsix` file in the current directory.

> You can also zip the contents of `vss-extension.json` and the `dist/` directory yourself; just rename the archive to
> `*.vsix`.

## Uploading to Azure DevOps Server (on-premises)

1. Sign in to your Azure DevOps Server (for this deployment: `https://azure.buluttakin.com`).
2. Open **Organization settings** (or **Collection settings** in Azure DevOps Server) → **Extensions** → **Manage extensions**.
3. Choose **Upload new extension**, select the generated `.vsix`, and upload it.
4. After upload, choose **Install** for the target project collection. The branch menu action will appear once installed.

> Tip for `azure.buluttakin.com`: if you are publishing from another machine, ensure the hostname is resolvable/reachable from
> that machine (for example via VPN or hosts file) before running any `tfx` command.

### Publishing with `tfx-cli` (alternative)

If you prefer the CLI, create a Personal Access Token (PAT) with the **Manage** extension permission. Then run:

```bash
tfx extension publish \
  --service-url https://azure.buluttakin.com \
  --token YOUR_PAT \
  --vsix <path-to-generated-file>.vsix
```

Use `--update` when pushing a new version, and increment `version` in `vss-extension.json` each time.

## Troubleshooting extension validation failures

If the Azure DevOps gallery reports `Error` or shows a validation message such as
`Something went wrong, please retry after sometime.`, use these steps to diagnose
the upload:

1. Capture the validation details from the gallery REST API. Replace
   `PUBLISHER`, `EXTENSIONNAME`, and `VERSION` with your values:

   ```bash
   curl -s \
     "https://azure.buluttakin.com/_apis/public/gallery/publisher/PUBLISHER/extension/EXTENSIONNAME/VERSION" \
     | jq .versions[0]
   ```

   Look for `validationResultMessage` and any missing `files` entries in the
   response. A message that stays generic after multiple retries usually means
   the package failed server-side validation (for example due to an unexpected
   manifest format or missing assets).

2. Verify that the VSIX contains both `Microsoft.VisualStudio.Services.VsixManifest`
   and `Microsoft.VisualStudio.Services.Manifest` assets. If you zipped the
   extension manually, ensure you included `vss-extension.json`, the entire
   `dist/` directory, and that the VSIX file name is unique for each version.

3. Confirm that the `version` field in `vss-extension.json` was incremented
   before packing. Azure DevOps rejects re-uploads with the same version even if
   the previous attempt failed.

4. For on-premises servers, double-check the `targets` version range. The
   manifest currently allows `[16.0,20.0)` (Azure DevOps Server 2019–2022).
   Older or newer servers may require adjusting this range to avoid
   `versionCheckError` validation failures.

5. After making fixes, regenerate the VSIX (`tfx extension create --rev-version`)
   and retry the upload. If the error persists, review the server event logs for
   extension validation errors—they often include the specific manifest or file
   issues that the public API hides behind the generic message.

## Troubleshooting access token errors on Azure DevOps Server

The extension depends on the host page to issue a scoped access token. On some
on-premises Azure DevOps Server deployments the token API can return HTTP 500
with `HostAuthorizationNotFound` even for collection administrators. To unblock
the generator:

1. Open **Organization/Collection settings** → **Extensions** → **Manage** and
   confirm the Pipeline Repo Generator extension is **installed** and **enabled**
   for the current collection and project. If it is disabled, enable it and
   retry the branch action.
2. If the extension shows as installed but still fails, click **Manage** on the
   extension details page and ensure the project you are using is selected.
   Removing and re-adding the project can refresh the host authorization entry
   that Azure DevOps uses to issue access tokens.
3. If the Manage view already shows your project, toggle the extension to
   disabled, re-enable it, then reload the Repos page. If that still fails,
   uninstall and reinstall the VSIX so Azure DevOps recreates the host
   authorization record for the collection/project.
4. After updating the extension permissions, reload the Repos page and relaunch
   the generator. If the error persists, sign out and back in to refresh the
   host session.

## Troubleshooting pipeline creation permission failures (401/TF400813)

If the form reports `Automatic pipeline creation failed: access was denied` or
the browser console shows `TF400813: The user is not authorized to access this
resource` when the generator tries to create the pipeline, the scoped token from
Azure DevOps likely does not include the **Create pipeline** permission for the
current project.

1. Ask a project administrator to grant your identity (or a security group you
   belong to) the **Create pipeline** permission under **Project settings** →
   **Pipelines** → **Security**. Retry after the permission change.
2. If you cannot obtain the permission immediately, you can still finish the
   process manually using the YAML that the generator pushed to the scaffold
   repository on the `main` branch. In Azure DevOps, go to **Pipelines** → **New
   pipeline** → **Azure Repos Git** → **Existing Azure Pipelines YAML**, then
   pick the `main` branch and the path shown in the generator status message
   (for example `/project-repo-env.yml`). You can also open the same screen
   directly (as in the screenshot) via
   `https://YOUR_SERVER/YOUR_COLLECTION/YOUR_PROJECT/_build?view=pipelines`,
   then choose **Existing Azure Pipelines YAML** in the right-hand panel and set
   the branch/path.
3. To verify your credentials outside the extension, call the same REST
   endpoint the generator uses:

   ```bash
  curl -u :<PAT_WITH_BUILD_SCOPE> \
    -H "Content-Type: application/json" \
    -d @pipeline.json \
    "https://YOUR_SERVER/YOUR_COLLECTION/YOUR_PROJECT/_apis/pipelines?repositoryId=<REPO_GUID>&api-version=7.1-preview.1"
   ```

   Replace the placeholders with your server, collection, and project, and
   include the pipeline payload in `pipeline.json` (for example the `name` and
   `configuration` object the generator attempted to send). A 401/TF400813
   response here confirms the token still lacks pipeline creation rights.

   A minimal request body looks like this (omit secrets and adjust the
   repository/path values for your project). Leaving `pipeline.json` empty will
   return `Value cannot be null. Parameter name: inputParameters` because the
   API expects these fields. The `repository.id` **must** be the GUID of the
   repository that contains your YAML (for example `HRMS_Azure_DevOps`), and the
   REST URL itself should include that repository ID as a `repositoryId`
   parameter. You can copy the GUID from **Repos** → **Files** → **Clone** (URI
   contains `.../_git/<repoId>`) or by calling `/_apis/git/repositories?api-version=6.0`.

   ```json
   {
     "name": "HRMS_hrms_demo",
     "configuration": {
       "type": "yaml",
       "path": "/HRMS-HRMS_Azure_DevOps-main.yml",
       "repository": {
         "id": "<REPO_GUID>",
         "type": "azureReposGit",
         "name": "BulutCollection/HRMS/_git/HRMS_Azure_DevOps",
         "defaultBranch": "refs/heads/main"
       }
     }
   }
   ```

## Local service hook testing (on-premises friendly)

The Azure DevOps service hook samples (see [official docs](https://learn.microsoft.com/azure/devops/extend/develop/add-service-hook))
recommend validating your webhook endpoint locally before wiring it to your organization or collection. This repository now
includes a minimal listener to mimic that flow and to keep your extension compatible with on-premises Azure DevOps Server.

1. Start the local listener (defaults to port `3000`):

   ```bash
   npm run service-hook:listen
   ```

   - Override the port with `npm run service-hook:listen -- --port 8081` (or `-p 8081`).
   - Disable raw payload logging with `npm run service-hook:listen -- --quiet` or by setting
     `LOG_PAYLOADS=false`.

   The listener logs the `eventType`, `notificationId`, collection, project, and repository values for every POST payload and
   always replies with HTTP 200 so Azure DevOps sees the connection as healthy.

2. If your Azure DevOps Server cannot reach `localhost`, expose the listener with a tunnel such as `ssh -R`, `Cloudflared`, or
   `ngrok`, and use the public URL in the next step. For the `azure.buluttakin.com` server, this makes it easy to send test
   payloads from your collection to a laptop listener.

3. In Azure DevOps (either Services or Server), open **Project settings** → **Service hooks** → **Create subscription** and pick
   **Web Hooks**. Use the listener URL (for example `http://localhost:3000/` or your tunnel URL) and click **Test** to send a
   sample payload. When the Azure DevOps Server at `https://azure.buluttakin.com` sends the request you will see the remote
   address in the listener output, confirming connectivity from that host.

4. Observe the console output from the listener to verify the payload shape before you depend on it in your extension or other
   downstream tooling.

5. When you are satisfied, package the extension with `tfx extension create --manifest-globs vss-extension.json --rev-version`
   and upload it to your on-premises collection as described above. Because the listener uses only Azure DevOps-standard fields
   (`eventType`, `notificationId`, `resourceContainers`, etc.), the same payload contract will be honored after deployment.

## Automatic provisioning from the extension

The extension now completes the workflow from the **Generate pipeline** button. After it writes the YAML to
`<ProjectName>_Azure_DevOps` on `main`, it runs these user-visible steps:

1. Create/reuse the generated repository.
2. Save/update the generated YAML.
3. Set `main` as that repository's default branch.
4. Create/update the YAML pipeline under `\komodo`, linked to that exact YAML path and `refs/heads/main`.
5. Create/reuse a classic Release definition named `<PipelineName>_Release` in `\komodo`. Its primary **Build** artifact is the pipeline from step 4 and its single agent job runs `Bash@3` using the selected pool and an inline script.

The UI stays on the form until all five stages complete, shows the precise failed stage, and then opens the generated Pipeline—not merely the YAML file. A failure in Release creation leaves the YAML pipeline available; correct the permission/configuration issue and run the generator again. Existing pipeline/release definitions are reused by name.

### Configure the Release Bash job

Edit `dist/release-config.js` **before packaging**. This file is included in the VSIX and must never contain a PAT, password, client secret, or other credential.

To keep the Bash code directly in the generated Release definition, replace the `content` value:

```js
scriptSource: {
  type: 'inline',
  content: `#!/usr/bin/env bash
set -euo pipefail

echo "Deploying $SERVICE_NAME"
# Your predefined Komodo commands go here.
`
}
```

To source the inline text from a file in a **same-collection Azure Repos** repository at generation time, replace `scriptSource` with the following. The running user needs read permission to that project/repository, and the retrieved text is still stored as the Release task's inline script:

```js
scriptSource: {
  type: 'azureReposFile',
  project: 'Tools',
  repository: 'deployment-scripts',
  branch: 'main',
  path: '/komodo/release-task.sh'
}
```

For the automatic workflow, the selected **Pool** must also be available as a project agent queue and the user must have permission to use it in Releases. The extension does not automatically create a release *instance*; it creates the reusable Release **definition**, as requested.

### Required Azure DevOps permissions for people using the button

The extension manifest requests `vso.code`, `vso.code_manage`, `vso.build`, `vso.build_execute`, `vso.release`, and `vso.release_manage`. These are requested scopes, not blanket authorization. A project administrator should grant users/groups only the permissions they need:

- **Repos:** Read/Contribute; additionally **Create repository** if the shared `<ProjectName>_Azure_DevOps` repository does not yet exist.
- **Pipelines:** View, **Create pipeline**, and Edit pipeline. The generated pipeline is created under `\komodo`.
- **Releases (Classic):** View releases and **Manage release definitions**. If your server has Release folder security, grant access to `\komodo`.
- **Agent queue:** **Use** permission on the queue selected in the form (for example `PublishDockerAgent`).
- **External script repository (optional):** Read access for the project/repository configured in `dist/release-config.js`.

If a user lacks one of these rights, Azure DevOps returns 401/403/TF400813 and the generator displays the failed stage and server response. This is intentional: the user cannot create a pipeline/release merely because an administrator published the extension.

### Do I need `AZP_TOKEN` while creating the extension?

**No.** Do not put `AZP_TOKEN` in `vss-extension.json`, JavaScript, HTML, VSIX packaging variables, or `release-config.js`. At runtime the extension calls `VSS.getAccessToken()` and sends the returned short-lived token as `Authorization: Bearer ...`; Azure DevOps issues it for the current signed-in user and enforces that user's permissions.

`AZP_TOKEN`/PAT is only needed outside the browser in these two administrative cases:

1. running `scripts/provision-pipeline-release.sh` from a terminal (legacy/automation path), or
2. publishing the completed VSIX through `tfx-cli`.

Use a separate, least-privileged **extension-management PAT** for publishing. Never reuse it as an application runtime token.

## Terminal provisioning fallback

After the YAML file has been pushed to the target repository, run the provisioning script to create:

1. a YAML pipeline connected to that file under the `komodo` pipeline folder, and
2. a classic Release definition that uses the created pipeline as its Build artifact and contains one inline `Bash@3` task.

Example:

```bash
export AZP_TOKEN="YOUR_PAT_OR_SCOPED_TOKEN"
export ADO_URL="https://azure.buluttakin.com"
export COLLECTION="ShonizCollection"
export PROJECT="Locanit"
export PIPELINE_NAME="Locanit_QA_Tester_qa"
export REPO_NAME="Locanit_QA"
export YAML_PATH="/qa/pipeline.yml"

# Required for the classic release agent job. Use either an ID or a name.
export RELEASE_AGENT_QUEUE_NAME="PublishDockerAgent"

# Your predefined inline Bash code can be read from a local file...
export RELEASE_BASH_SCRIPT_FILE="./scripts/release-inline-task.example.sh"

npm run pipeline:provision-release
```

The script is idempotent by name: if the pipeline or release definition already exists, it reuses the existing ID. Defaults:

- `PIPELINE_FOLDER=komodo` creates the pipeline under `\komodo`.
- `RELEASE_NAME=${PIPELINE_NAME}_Release`.
- `RELEASE_FOLDER=komodo` creates the release definition under `\komodo`.
- `RELEASE_ENVIRONMENT_NAME=komodo`.
- `DEFAULT_BRANCH=refs/heads/main`.

To load the Bash task from another repository instead of a local file:

```bash
export RELEASE_BASH_SCRIPT_GIT_URL="https://azure.buluttakin.com/ShonizCollection/Tools/_git/deployment-scripts"
export RELEASE_BASH_SCRIPT_GIT_REF="main"
export RELEASE_BASH_SCRIPT_GIT_PATH="komodo/release-task.sh"
npm run pipeline:provision-release
```

If you also want to create an actual Release run immediately after creating/reusing the definition, set:

```bash
export CREATE_RELEASE_INSTANCE=true
```

### Error handling and diagnostics

`scripts/provision-pipeline-release.sh` prints a `[STEP]` line before each major action and returns clear `[ERROR]` messages for
the failed step. REST API errors include the method, URL, HTTP status code, and response body so permission, endpoint, and payload
issues can be diagnosed without re-running with extra debug flags.

### If the generator only created the YAML file and redirected to it

Older extension builds could push the YAML file and immediately redirect to the file page before registering the Azure Pipeline.
Install version `0.1.17` or later, then run **Generate pipeline** again for the same branch. The generator now:

1. creates or updates the YAML file in `<ProjectName>_Azure_DevOps`,
2. creates or updates the Azure Pipeline under `\komodo`, and
3. redirects only after the pipeline API returns successfully.

If pipeline registration fails, the form stays open and shows the REST/API error instead of redirecting. Check that the user has
**Create pipeline** and **Edit pipeline** permissions in **Project settings → Pipelines → Security**.

## Build and upload the extension to on-premises Azure DevOps Server

The extension can be packaged with `tfx-cli`. You can run it through `npx` without installing a global package.

> Before packaging, make sure every path referenced by `vss-extension.json` exists in the repository. In this manifest that means
> the `dist/` folder, `dist/images/icon.svg`, and `README.md` must be present. If `dist/` is generated or copied from another
> source, place it in the repo root before running the package command.

1. Update the release Bash configuration in `dist/release-config.js`, and validate all local assets:

   ```bash
   npm test
   ```

2. Package the VSIX (this uses `npx`; no global installation is required):

   ```bash
   npm run extension:package
   ```

   Equivalent direct command:

   ```bash
   npx --yes tfx-cli extension create \
     --manifest-globs vss-extension.json \
     --rev-version
   ```

3. Locate the package and upload/publish the generated VSIX to your on-premises Azure DevOps Server gallery. Use a separate PAT that has **extension management** permission only:

   ```bash
    export ADO_SERVICE_URL="https://azure.buluttakin.com"
    export AZP_TOKEN="YOUR_EXTENSION_MANAGEMENT_PAT"
    export VSIX_FILE="$(find . -maxdepth 1 -type f -name 'mohammad-falahat.pipeline-generator-*.vsix' -print -quit)"
    test -n "$VSIX_FILE" || { echo "VSIX was not created" >&2; exit 1; }

   npx --yes tfx-cli extension publish \
     --service-url "$ADO_SERVICE_URL" \
     --token "$AZP_TOKEN" \
     --vsix $VSIX_FILE
   ```

    Or with the npm script (the package file is passed after `--`):

   ```bash
   export ADO_SERVICE_URL="https://azure.buluttakin.com"
   export AZP_TOKEN="YOUR_EXTENSION_MANAGEMENT_PAT"
    npm run extension:publish -- "$VSIX_FILE"
   ```

4. If your server requires manual installation after upload, open:

   ```text
   https://azure.buluttakin.com/_gallery/manage
   ```

    Then install/enable the uploaded extension for the target collection/project. Reload the Repos page (hard refresh if needed), then use **Generate pipeline** on a test branch.

### Copy/paste build and upload commands for `azure.buluttakin.com`

```bash
cd /home/falahat/azure-devops-pipeline-repo-generator

# 1) Edit dist/release-config.js with the real Bash content or Azure Repos source.
# 2) Run static/offline checks. No token is needed here.
npm test

# 3) Generate a versioned VSIX. --rev-version increments the manifest version.
npm run extension:package

# 4) Publish only with an extension-management PAT; this PAT is NOT bundled in the VSIX.
export ADO_SERVICE_URL='https://azure.buluttakin.com'
read -rsp 'Extension-management PAT: ' AZP_TOKEN; echo
export AZP_TOKEN
export VSIX_FILE="$(find . -maxdepth 1 -type f -name 'mohammad-falahat.pipeline-generator-*.vsix' -print -quit)"
test -n "$VSIX_FILE" || { echo 'VSIX was not created' >&2; exit 1; }
npx --yes tfx-cli extension publish \
  --service-url "$ADO_SERVICE_URL" \
  --token "$AZP_TOKEN" \
  --vsix "$VSIX_FILE"

# 5) In Collection Settings → Extensions, install/enable it for ShonizCollection,
#    then test it in a project such as Area or Locanit.
unset AZP_TOKEN
```
