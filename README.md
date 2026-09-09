# Azure DevOps Pipeline Repo Generator

Pipeline Generator is a static Azure DevOps extension that adds **Generate
pipeline** and **Generate MonoRepo** actions to Azure Repos branch menus. From a selected source branch it
creates or reuses a project-level repository, writes a generated YAML file,
registers a YAML Pipeline that points to that file, and creates a classic
Release definition that consumes the Pipeline as a Build artifact.

The manifest version documented here is **0.1.67**. The manifest targets Azure
DevOps Services and Azure DevOps Server range `[16.0,20.0)`. The complete live
workflow has been verified on the documented on-premises Server environment;
the Azure DevOps Services Release API route still requires a separate
compatibility test.

## What the program creates

```text
selected source repository + branch
        |
        | reads environments from ShonizCollection/SharedTemplates
        | reads central Komodo credentials from ShonizCollection/SharedTemplates
        | reads enabled servers directly from Komodo
        | resolves nginx-net or nginx-network on the selected Docker host
        | ensures <ProjectName>_Docker_DevOps and <ProjectName>_Nginx_DevOps
        |
        | generates a Branch-to-Environment-specific YAML document
        v
<ProjectName>_Azure_DevOps @ main
        |
        | YAML configuration points to this repository and file
        v
YAML Pipeline in \komodo
        | name: <project>-<repository>-<Branch>To<ENVIRONMENT>.yml
        |
        | primary Build artifact
        v
Classic Release definition in \komodo
        | name: <SERVICE> <ENVIRONMENT> (for example API DEMO)
        | links project Variable Group KomodoAPI
        |
        v
one agent-based Bash@3 deployment job with packaged wrapper stored Inline
```

## Nx Monorepo mode

**Generate MonoRepo** creates a separate, non-conflicting MR path while leaving
the normal generator unchanged:

- one Pipeline named
  `<project>-<repository>-MR-<service>-<Branch>To<ENVIRONMENT>.yml` under `\komodo\MR`;
- one classic Release named `MR <SERVICE> <ENVIRONMENT>` under `\komodo\MR`;
- an automatically created `/.devops/deployments.yml` project contract, with
  the shared `monorepo/pipeline.yml` and `monorepo/mr-build.cjs` loaded from
  `ShonizCollection/SharedTemplates`; the same template packages the central
  `monorepo/nginx/default.conf` instead of generating it inside Compose;
- one logical Monorepo service, represented by a generic Nginx static runtime
  image and an optional Node BFF image, merged into the existing project and
  Environment `compose.yml`. That shared file remains on `main` in ADO and is
  the GitOps source of truth. Generic Dockerfiles are owned by SharedTemplates,
  so the source project does not need a Dockerfile;
- `/bff/` routing to the Monorepo BFF, leaving `/api/` reserved for the main
  backend, and root routing to the shell/static runtime,
  without URI rewrite and with the root Location last;
- one `mr-drop` artifact containing a full module inventory plus only the
  affected build outputs. A shell/host change rebuilds every buildable Nx app;
  other changes build only affected applications. Each module is built
  independently: a failed ordinary module retains its previous deployed
  version while successful modules continue; a failed shell blocks deployment;
- immutable versioned images in the selected internal registry. The first run
  performs a full Nx build. Later runs hydrate the previously active static
  image, overlay successful affected outputs, retain failed/unaffected outputs,
  and push a new tag. The BFF is rebuilt only when affected and otherwise keeps
  its prior tag. The central Pipeline template creates or updates a Komodo Repo linked to the
  ADO Docker repository and partially reconciles the same project/Environment
  Stack used by ordinary services. Existing Stack settings and existing
  Compose services are retained. Stable Compose image fields reference
  service-specific variables such as `${front_monorepo}` and
  `${front_monorepo_bff}`; the adjacent tracked `.env` owns the immutable tags.
  The Release updates only those `.env` values after a one-time migration of
  legacy hard-coded image fields, asks Komodo to `DeployStack`, and polls the returned
  Update. On failure it restores the exact previous Compose/`.env` state and redeploys it. No Komodo
  Terminal access, target-side artifact staging, host bind mount, Node, or
  Docker is required on the Release agent.
- bounded Komodo busy handling: `CreateStack`, `UpdateStack`, and `DeployStack`
  retry up to five total attempts with a five-second interval only when Komodo
  reports `Stack busy`; permanent errors still fail immediately.

Before Step 1 writes any support file, the generator calls Komodo
`ListDockerNetworks` for the selected Server. It accepts only an exact existing
`nginx-network` or `nginx-net`. Existing Compose service references are
preserved through an explicit external-network `name` mapping, so a repository
that uses the logical key `nginx-network` can safely target a host whose actual
network is `nginx-net`. Missing or ambiguous unsupported network names stop
provisioning before repository writes.

New buildable Nx applications are discovered on the next run. A rename is
treated as a new application plus an orphaned old application: the new output
is activated, while the old output is retained and reported instead of being
deleted. The generated contract is created only when missing, so later manual
command/name overrides are preserved.
Static module outputs are linked under `/<nx-project-name>/` inside the immutable
runtime, while the shell remains at `/` and Monorepo BFF traffic remains at
`/bff/`. The main application backend retains `/api/`.

The central `komodo-servers-creds.env` key remains read-only and is used only
to populate the Server select. The separate `KomodoAPI` Variable Group
credentials used by MR Build/Release need permission to create/update Komodo
Repo and Stack resources and execute `DeployStack` on the selected Server;
they are not read from or embedded in the browser bundle.

The browser workflow performs five visible steps:

1. Resolve the selected Server's external Nginx Docker network, then create or reuse `<ProjectName>_Azure_DevOps`,
   `<ProjectNameWithoutSpaces>_Docker_DevOps`, and
   `<ProjectNameWithoutSpaces>_Nginx_DevOps`; initialize the two support
   repositories with a starter `compose.yml`, its adjacent tracked `.env`, and one shared
   project/Environment Nginx configuration with managed service routes.
2. Add or edit the generated YAML file on `main`.
3. set `refs/heads/main` as the generated repository's default branch.
4. Create or update the YAML Pipeline under `\komodo`.
5. Create or reconcile the classic Release definition under `\komodo`, embed
   the packaged Bash wrapper Inline, and link Variable Group `KomodoAPI`.

After all enabled steps succeed, the page stays open, collapses and locks the
completed form, and moves focus to links for the starter Nginx file, starter
`compose.yml`, and real Pipeline. The user reviews and edits both files before
opening and manually running the Pipeline. The extension does not queue a
Pipeline run and does not create a Release instance.

## Important concepts

- The **source repository** is the repository from whose branch menu the user
  launches the extension.
- The **central configuration repository** is always
  `ShonizCollection/SharedTemplates/SharedTemplates`, even when the selected
  source project belongs to another collection. The signed-in extension user
  must have Read permission there.
- The **source branch** is embedded in the generated YAML as the `otherRepo`
  repository resource.
- The **generated repository** is the shared project repository named
  `<ProjectName>_Azure_DevOps`.
- The **Docker DevOps repository** is named
  `<ProjectNameWithoutSpaces>_Docker_DevOps` and receives
  `<environment>_<lowercase-project-without-spaces>/compose.yml`; Monorepo mode
  also ensures the adjacent `.env` used for managed immutable image tags. Its starter
  service/container is `<project>_<service>_<environment>` and exposes port 80
  for UI/frontend services or 8080 for backend and other services. Its external Nginx
  network is resolved from the selected Komodo Server and may be either
  `nginx-network` or `nginx-net`; the Compose network's explicit `name` maps
  preserved logical service references to the actual host network.
- The **Nginx DevOps repository** is named
  `<ProjectNameWithoutSpaces>_Nginx_DevOps` and receives
  `<environment>/<lowercase-project>-<environment>.conf`. Its host is
  `<lowercase-sanitized-project>.<environment-domain>`. UI/frontend indicators
  (`ui`, `front`, `frontend`, `web`, `client`, and similar aliases) use `/`,
  while backend/API indicators (`api`, `back`, `backend`, `server`, and similar
  aliases) use `/api/`; other services use `/<service>/`. A recognized newer
  variant such as `UI_V2`, `BACK_v2`, `NewUI`, or `api_refactor` is isolated at
  `/v2/`, `/api/v2/`, `/new/`, or `/api/refactor/` respectively. A repository
  whose name exactly matches its project is always treated as the preferred
  frontend owner of `/`. Before assigning `/` or `/api/`, the generator reads
  all project repositories and scores every frontend/backend candidate by
  semantic specificity. A pure role name beats one with extra qualifiers, and
  fewer/shorter qualifiers beat broader names; no repository-name pair is
  predefined. For example, `front` owns `/` ahead of `front_panel`, whose route
  becomes `/front-panel/`. Separated and joined forms such as `front_panel` and
  `frontpanel` are both recognized. This decision is independent of Pipeline execution order.
  Each later service run reads
  the shared file and inserts only a missing direct-child Location into the
  managed-routes section. An existing Location and all manual content are
  preserved; ambiguous duplicate HTTPS server blocks stop automatic editing.
  Every managed route uses Docker DNS (`resolver 127.0.0.11 ipv6=off`) and
  stores its container hostname in `$target`. Root proxies to
  `http://$target:80`. Non-root `/<service>/` routes proxy to
  `http://$target:8080` without a URI slash or rewrite, preserving the original
  request URI. The root Location is always ordered below all other managed
  Locations. Older managed paths, generated rewrites, and proxy forms are
  migrated automatically. Exact legacy generated certificate paths such as
  `bulutdemo.pem` and `bulutdemo.key` are also migrated to the complete
  Environment domain (`bulutdemo.ir.pem` and `bulutdemo.ir.key`) inside only
  the matching HTTPS server block; custom certificate paths remain unchanged.
- The **scaffold branch** is always `main`; generated YAML files and Pipeline
  definitions point to this branch.
- The Pipeline display name is exactly the generated YAML filename, including
  its `.yml` suffix. This makes the file/Pipeline relationship visible and
  keeps different source branches independently named.
- The classic Release display name contains only the uppercased Service and
  Environment form values, for example `UI DEMO` or `API PRO`.

## Documentation map

| Document | Use it for |
| --- | --- |
| [دستورالعمل استقرار سرویس مونوریپو](docs/monorepo-service-deployment-fa.md) | آماده‌سازی Nx، Dockerfileهای مرکزی، Compose بدون bind mount و مراحل Build/Release |
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
| `dist/ui.js` | Form hydration, starter-file/YAML rendering, Azure DevOps REST orchestration, error handling, and completion links |
| `dist/release-config.js` | Token-free classic Release settings, required Variable Group, and packaged Bash source selection |
| `dist/release-inline-task.sh` | Bash wrapper embedded as inline task text in each generated/reconciled Release definition |
| `dist/monorepo-build.cjs` | Maintained mirror of the central SharedTemplates runner that discovers Nx applications, computes affected projects, builds them, and writes `mr-drop` metadata/output |
| `dist/monorepo-release-inline-task.sh` | Release task that commits immutable image tags to the tracked Compose `.env`, migrates legacy image fields, deploys the Komodo Stack, polls its Update, and restores the prior Compose/`.env` state on failure |
| `dist/lib/VSS.SDK*.js` | Bundled legacy VSS Web Extension SDK used by the on-premises host |
| `dist/styles.css` | Generator page styling |
| `scripts/validate-extension.js` | Offline static contract validation |
| `scripts/validate-action-behavior.js` | Mocked regression test for Dialog/Hub launch and token-free hosted context |
| `scripts/validate-ui-behavior.js` | Mocked REST regression tests for exact naming and non-destructive Pipeline/Release migration |
| `scripts/provision-pipeline-release.sh` | Terminal fallback for provisioning a Pipeline and classic Release after YAML already exists |
| `scripts/service-hook-listener.js` | Development-only listener for inspecting Azure DevOps Service Hook payloads |
| `scripts/release-inline-task.example.sh` | Example Bash source for the normal shell provisioner |

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

The browser extension persists no credential. It reads the centrally managed
read-only Komodo credential into page memory only while loading enabled server
names, then discards the local function scope. The branch action first asks
Azure DevOps to open `dist/index.html` through
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

Never hard-code a PAT, password, cookie, registry credential, or API key in
JavaScript, HTML, the manifest, `release-config.js`, documentation, or a VSIX.
By explicit operator policy, the dedicated Server-Read Komodo credential is a
browser-readable central Azure Repos file; the runtime never logs or persists
it. The generated YAML and classic Release continue to reference their separate
Komodo deployment credentials through the `KomodoAPI` Variable Group.

## Configuration summary

Pool and registry defaults are declared in `dist/ui.js`. Environment names and
domains are read at runtime from
`ShonizCollection/SharedTemplates/SharedTemplates:/pipeline-generator.yml` on
`main`, using the signed-in browser session. The extension deliberately omits
the current collection's Bearer token for these same-origin cross-collection
reads because Azure DevOps Server rejects that token in a sibling collection.
The preferred contract is:

```yaml
environments:
  - name: dev
    domain: bulutdev.ir
```

Every environment requires a valid domain; the compact legacy value
`"dev:bulutdev.ir"` is accepted for migration. Any legacy `servers:` list is
ignored by the form. To load Komodo Server choices, the extension reads
`ShonizCollection/SharedTemplates/SharedTemplates:/komodo-servers-creds.env@main`,
calls Komodo
1.19.x `ListFullServers` directly, filters strictly on
`config.enabled === true`, excludes templates, and retains only names. Release
definition behavior is declared in `dist/release-config.js`.

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
