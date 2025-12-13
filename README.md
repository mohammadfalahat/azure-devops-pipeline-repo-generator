# Azure DevOps Pipeline Repo Generator

A lightweight Azure DevOps extension that adds a **Generate pipeline** action beside each branch. When invoked, it opens a form pre-populated with the required deployment settings:

- `pool`: `PublishDockerAgent`
- `service`: `api`
- `environment`: `demo`
- `dockerfileDir`: `src/TMS.API`
- `repositoryAddress`: `registry.buluttakin.com`
- `containerRegistryService`: `BulutReg`
- `komodoServer`: `DEMO-192.168.62.91` (with other options available)

Submitting the form ensures a shared repository named `SANITIZEDPROJECTNAME_Azure_DevOps` exists in the current project. If it does not, the extension creates it and pushes a `pipeline-template.yml` containing the submitted settings.

## Structure

- `vss-extension.json`: Extension manifest defining the branch action and a hub entry.
- `src/index.html`: Form UI loaded from the action or the hub.
- `src/ui.js`: Client-side logic that sanitizes the project name, creates the repository, and pushes the template.
- `src/menu-action.js`: Registers the branch menu action and opens the UI with branch context.
- `src/styles.css`: Lightweight styling for the form.

## Packaging (creating the VSIX)

The extension is static (HTML/CSS/JS only), so packaging is just zipping the manifest and `src` folder into a VSIX. You can do
this manually or with the official `tfx-cli` utility.

### Prerequisites

- Node.js 18+ (for `npm` and `tfx-cli`).
- A publisher ID and display name to embed in `vss-extension.json` (`publisher` and `name` fields). For on-premises servers you
  can use any unique publisher ID (it is not tied to the public marketplace).
- An icon at `src/images/icon.svg` (a 128x128 SVG) if you want to replace the placeholder.

### Using `tfx-cli`

1. Install the CLI: `npm install -g tfx-cli` (or download the release ZIP if your server blocks npm registry traffic).
2. Bump the `version` in `vss-extension.json` as needed.
3. Run the pack command from the repo root:

   ```bash
   tfx extension create --manifest-globs vss-extension.json --rev-version
   ```

   This outputs a `*.vsix` file in the current directory.

> You can also zip the contents of `vss-extension.json` and the `src/` directory yourself; just rename the archive to
> `*.vsix`.

## Uploading to Azure DevOps Server (on-premises)

1. Sign in to your Azure DevOps Server (for example `https://azure.buluttakin.com/tfs`).
2. Open **Organization settings** (or **Collection settings** in Azure DevOps Server) → **Extensions** → **Manage extensions**.
3. Choose **Upload new extension**, select the generated `.vsix`, and upload it.
4. After upload, choose **Install** for the target project collection. The branch menu action will appear once installed.

### Publishing with `tfx-cli` (alternative)

If you prefer the CLI, create a Personal Access Token (PAT) with the **Manage** extension permission. Then run:

```bash
tfx extension publish \
  --service-url https://azure.buluttakin.com/tfs \
  --token YOUR_PAT \
  --vsix <path-to-generated-file>.vsix
```

Use `--update` when pushing a new version, and increment `version` in `vss-extension.json` each time.
