# Azure DevOps Pipeline Repo Generator

Pipeline Generator is a static Azure DevOps extension that adds a **Generate
pipeline** action to Azure Repos branch menus. From a selected source branch it
creates or reuses a project-level repository, writes a generated YAML file,
registers a YAML Pipeline that points to that file, and creates a classic
Release definition that consumes the Pipeline as a Build artifact.

The manifest version documented here is **0.1.32**. The manifest targets Azure
DevOps Services and Azure DevOps Server range `[16.0,20.0)`. The complete live
workflow has been verified on the documented on-premises Server environment;
the Azure DevOps Services Release API route still requires a separate
compatibility test.

## What the program creates

```text
selected source repository + branch
        |
        | generates a branch-specific YAML document
        v
<ProjectName>_Azure_DevOps @ main
        |
        | YAML configuration points to this repository and file
        v
YAML Pipeline in \komodo
        |
        | primary Build artifact
        v
Classic Release definition in \komodo
        | links project Variable Group KomodoAPI
        |
        v
one agent-based Bash@3 deployment job with packaged wrapper stored Inline
```

The browser workflow performs five visible steps:

1. Create or reuse `<ProjectName>_Azure_DevOps`.
2. Add or edit the generated YAML file on `main`.
3. set `refs/heads/main` as the generated repository's default branch.
4. Create or update the YAML Pipeline under `\komodo`.
5. Create or reconcile the classic Release definition under `\komodo`, embed
   the packaged Bash wrapper Inline, and link Variable Group `KomodoAPI`.

After all enabled steps succeed, the page redirects to the real Pipeline. It
does not redirect merely because the YAML file was created. The extension does
not queue a Pipeline run and does not create a Release instance.

## Important concepts

- The **source repository** is the repository from whose branch menu the user
  launches the extension.
- The **source branch** is embedded in the generated YAML as the `otherRepo`
  repository resource.
- The **generated repository** is the shared project repository named
  `<ProjectName>_Azure_DevOps`.
- The **scaffold branch** is always `main`; generated YAML files and Pipeline
  definitions point to this branch.
- The Pipeline display name is exactly the generated YAML filename, including
  its `.yml` suffix. This makes the file/Pipeline relationship visible and
  keeps different source branches independently named.

## Documentation map

| Document | Use it for |
| --- | --- |
| [Architecture and runtime flow](docs/architecture.md) | Components, state transfer, generated YAML, provisioning sequence, reconciliation, naming, and design constraints |
| [Azure DevOps REST contracts](docs/rest-api-contracts.md) | Authentication, scopes, permissions, endpoints, API versions, payload invariants, and on-premises behavior |
| [Development and operations](docs/development-and-operations.md) | Local setup, configuration, testing, packaging, installation, shell automation, release procedure, and troubleshooting |
| [Verified E2E context](docs/azure-devops-e2e-context.md) | Layered findings and forensic history from the real on-premises environment |
| [E2E state snapshot](docs/azure-devops-e2e-state.yaml) | Machine-readable IDs, versions, links, and last verified state |

Maintainers and coding agents must read the E2E context and state snapshot
before repeating live provisioning tests. They contain resources that must not
be deleted during routine cleanup.

## Repository layout

| Path | Responsibility |
| --- | --- |
| `vss-extension.json` | Extension identity, version, branch action, Dialog and Azure Repos Hub contributions, host compatibility, requested scopes, and packaged assets |
| `dist/menu-action.html` | Minimal host page for the branch action contribution |
| `dist/menu-action.js` | Registers the action, extracts branch context, and opens a host Dialog or navigates to the in-host Azure Repos Hub |
| `dist/index.html` | Generator form hosted by both the Dialog control and Azure Repos Hub contributions |
| `dist/ui.js` | Form hydration, YAML rendering, Azure DevOps REST orchestration, error handling, and redirect |
| `dist/release-config.js` | Token-free classic Release settings, required Variable Group, and packaged Bash source selection |
| `dist/release-inline-task.sh` | Bash wrapper embedded as inline task text in each generated/reconciled Release definition |
| `dist/lib/VSS.SDK*.js` | Bundled legacy VSS Web Extension SDK used by the on-premises host |
| `dist/styles.css` | Generator page styling |
| `scripts/validate-extension.js` | Offline static contract validation |
| `scripts/validate-action-behavior.js` | Mocked regression test for Dialog/Hub launch and token-free hosted context |
| `scripts/validate-ui-behavior.js` | Mocked REST regression tests for exact naming and non-destructive Pipeline/Release migration |
| `scripts/provision-pipeline-release.sh` | Terminal fallback for provisioning a Pipeline and classic Release after YAML already exists |
| `scripts/service-hook-listener.js` | Development-only listener for inspecting Azure DevOps Service Hook payloads |
| `scripts/release-inline-task.example.sh` | Example Bash source for the shell provisioner |

There is no compilation or bundling stage. The files in `dist/` are the actual
runtime assets packaged into the VSIX.

## Developer quick start

Requirements: Node.js 18 or newer and npm. The shell automation additionally
requires Bash, `curl`, and `jq`; its Git script-source mode also requires Git.

```bash
npm ci
npm test
npm run build
```

`npm run build` intentionally performs no transformation. Before packaging,
configure the Release task in `dist/release-config.js` and rerun `npm test`.

Create a VSIX with:

```bash
npm run extension:package
```

The packaging script invokes `tfx-cli extension create --rev-version`. Inspect
the resulting manifest/version changes and the generated VSIX before publishing.
See [Development and operations](docs/development-and-operations.md) for the
complete release procedure.

## Runtime authentication and permissions

The browser extension has no backend and persists no credential. The branch
action first asks Azure DevOps to open `dist/index.html` through
`openCustomDialog`. If an older Azure DevOps Server does not expose that
service, the action navigates the current host page to the extension's
`pipeline-generator-hub` under Azure Repos. It never opens the Generator in a
new tab. Dialog context is passed through `VSS.getConfiguration()`; Hub context
is read from host navigation state. After its own iframe SDK handshake, the
Generator calls `VSS.getAccessToken()` and sends that short-lived current-user
token as a Bearer credential. The action neither obtains nor transfers a token.

If an on-premises host cannot issue the token (for example,
`HostAuthorizationNotFound`), the form offers an explicit **Sign out and
authenticate again** action plus a direct **Open extension authorization**
action. `HostAuthorizationNotFound` requires a Collection Administrator to
authorize the requested extension scopes in Collection Settings → Extensions;
signing in again cannot create missing extension authorization by itself. The
browser extension never asks for or accepts a PAT.

The manifest requests these scopes:

- `vso.code` and `vso.code_manage`
- `vso.project`
- `vso.build` and `vso.build_execute`
- `vso.release` and `vso.release_manage`
- `vso.agentpools` to resolve the selected Release agent queue
- `vso.serviceendpoint` to list Docker Registry service connections
- `vso.variablegroups_read` to resolve and link `KomodoAPI` to the Release

Scopes do not grant project permissions by themselves. The user also needs
repository Read/Contribute and possibly Create repository, Pipeline
Create/Edit, classic Release definition management, and Use permission on the
selected agent queue. Details are in
[Authentication and authorization](docs/rest-api-contracts.md#authentication-and-authorization).

Never hard-code a PAT, password, cookie, service secret, registry credential,
or API key in JavaScript, HTML, the manifest, `release-config.js`,
documentation, or a VSIX. The generated YAML and classic Release reference
Komodo credentials through the `KomodoAPI` Variable Group rather than embedding
their values.

## Configuration summary

The form defaults are declared in `dist/ui.js`. Release definition behavior is
declared in `dist/release-config.js`. The default deployment mapping is:

| Environment | Komodo target |
| --- | --- |
| `dev` | `Development-192.168.62.19` |
| `demo` | `DEMO-192.168.62.91` |
| `qa` | `QA-192.168.62.153` |
| `pro` | `Production-31.7.65.195` |

The generated YAML consumes `build-push-komodo.yml` from
`SharedTemplates/SharedTemplates` through the `ShonizCollection` service
endpoint and uses variable group `KomodoAPI`. These names are current
implementation contracts, not environment variables.

## Verification

Run the complete offline suite with:

```bash
npm test
```

It validates the manifest/runtime contract, mocked UI reconciliation behavior,
the Service Hook parser, and shell JSON payload construction. It does not
contact Azure DevOps and is not a replacement for a published-extension E2E
test. The last live result and its authentication boundary are recorded in the
E2E context documents.
