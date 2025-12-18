# Azure DevOps Pipeline Repo Generator

A lightweight Azure DevOps extension that adds a **Generate pipeline** action beside each branch. When invoked, it opens a form pre-populated with the required deployment settings:

- `pool`: `PublishDockerAgent`
- `service`: `api`
- `environment`: `demo`
- `dockerfileDir`: `src/TMS.API`
- `repositoryAddress`: `registry.buluttakin.com`
- `containerRegistryService`: `BulutReg`
- `komodoServer`: `DEMO-192.168.62.91` (with other options available)

The extension uses the scoped Azure DevOps access token provided by the host page to create repositories and scaffold the pipeline template—no extra prompts or saved tokens are required.

Submitting the form ensures a shared repository named `SANITIZEDPROJECTNAME_Azure_DevOps` exists in the current project. If it does not, the extension creates it and pushes a `pipeline-template.yml` containing the submitted settings.

## Structure

- `vss-extension.json`: Extension manifest defining the branch action and a hub entry.
- `dist/index.html`: Form UI loaded from the action or the hub.
- `dist/ui.js`: Client-side logic that sanitizes the project name, creates the repository, and pushes the template.
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
  loads the platform-hosted SDK from `/_content/MS.VSS.SDK/scripts/VSS.SDK.min.js` at runtime.
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

1. Sign in to your Azure DevOps Server (for example `https://azure.buluttakin.com/tfs`).
2. Open **Organization settings** (or **Collection settings** in Azure DevOps Server) → **Extensions** → **Manage extensions**.
3. Choose **Upload new extension**, select the generated `.vsix`, and upload it.
4. After upload, choose **Install** for the target project collection. The branch menu action will appear once installed.

> Tip for `azure.buluttakin.com`: if you are publishing from another machine, ensure the hostname is resolvable/reachable from
> that machine (for example via VPN or hosts file) before running any `tfx` command.

### Publishing with `tfx-cli` (alternative)

If you prefer the CLI, create a Personal Access Token (PAT) with the **Manage** extension permission. Then run:

```bash
tfx extension publish \
  --service-url https://azure.buluttakin.com/tfs \
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

## Local service hook testing (on-premises friendly)

The Azure DevOps service hook samples (see [official docs](https://learn.microsoft.com/azure/devops/extend/develop/add-service-hook))
recommend validating your webhook endpoint locally before wiring it to your organization or collection. This repository now
includes a minimal listener to mimic that flow and to keep your extension compatible with on-premises Azure DevOps Server.

1. Start the local listener (defaults to port `3000`):

   ```bash
   npm run service-hook:listen
   ```

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
