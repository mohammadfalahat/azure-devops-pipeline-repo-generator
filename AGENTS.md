# Repository working context

Before changing or retesting the Azure DevOps provisioning flow, read both of
these files:

- `docs/azure-devops-e2e-context.md` — layered findings, root causes, safe
  procedures, and approaches that already failed.
- `docs/azure-devops-e2e-state.yaml` — machine-readable snapshot of the last
  verified server state and generated resources.

For implementation work, also use the maintained technical documentation:

- `docs/architecture.md` — runtime components, state and provisioning flow.
- `docs/rest-api-contracts.md` — authentication and Azure DevOps REST contracts.
- `docs/development-and-operations.md` — configuration, tests, packaging,
  installation, live verification, and troubleshooting.

Keep the two files synchronized after any live end-to-end test. Clearly label
whether a result came from a normal signed-in Azure DevOps session, an
authentication-adapted browser test, a direct REST call, or local validation.

Never write a PAT, session token, cookie, authorization header, or other secret
to this repository, logs, command arguments, or generated evidence. The
on-premises Azure DevOps host is internal; bypass the environment HTTP proxy for
it. Do not delete the current successful verification resources unless the user
explicitly requests cleanup.
