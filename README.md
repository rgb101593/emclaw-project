# EMClaw

[![CI](https://github.com/rgb101593/emclaw-project/actions/workflows/ci.yml/badge.svg)](https://github.com/rgb101593/emclaw-project/actions/workflows/ci.yml)
[![Gitleaks](https://github.com/rgb101593/emclaw-project/actions/workflows/gitleaks.yml/badge.svg)](https://github.com/rgb101593/emclaw-project/actions/workflows/gitleaks.yml)

EMClaw is a Slack-based work assistant for Microsoft 365. It combines a local
language model with deterministic identity, authorization, and output checks.
The model handles conversation; ordinary code decides which user is calling,
which data may be read, and whether a response is safe to send.

I designed and built the connector, policy controls, local-model integration,
tests, and deployment safeguards. The main implementation is TypeScript, with a
small dependency-free Python walkthrough for the security model.

> **Sanitized public snapshot:** Credentials, private runtime state,
> organization-specific configuration, and live data are excluded. The demo and
> connector tests run locally. A complete Slack and Microsoft 365 deployment
> requires private configuration that is not included here.

## What it demonstrates

- Caller identity is derived from trusted runtime context rather than prompt text.
- Member actions are self-only; cross-user operations require an admin path.
- Microsoft Graph requests use restricted metadata fields.
- The deployment's outbound guards block unsupported success claims, unsafe
  connection claims, and credential-shaped output; the standalone demo shows
  the same fail-closed behavior without live services.
- Workflow termination moves state into quarantine and writes rollback records
  instead of deleting it immediately.

The language model runs locally. Slack and Microsoft Graph remain external
systems of record, and their data is not sent to a separate hosted LLM.

## Try it

The security walkthrough uses only the Python standard library (Python 3.9+):

```bash
make demo
```

Build and test the Microsoft 365 connector with Node.js 22.22.3+ or 24.15.0+:

```bash
make install
make typecheck
make test
```

You can also run the connector directly:

```bash
cd v2/plugins/emclaw-m365-connector
npm ci
npm run build
npm test
npm audit --omit=dev
```

For the complete connector check:

```bash
npm run verify
```

## Repository layout

```text
demo/                              Dependency-free security walkthrough
v2/plugins/emclaw-m365-connector/  TypeScript tool plugin and tests
  src/index.ts                     Tool registration and trusted-context adapters
  src/identity.ts                  Roster lookup and caller authorization
  src/calendar.ts                  Calendar queries and safe result shaping
  src/email.ts                     Email metadata queries and filtering
  src/preferences.ts               Per-user preference operations
  src/workflows.ts                 Workflow lifecycle and administration
  src/config.ts                    Runtime paths and policy settings
  src/schemas.ts                   Tool input schemas
  src/util.ts                      Shared filesystem and redaction helpers
v2/identity/                       Redacted roster data
```

See [ARCHITECTURE.md](ARCHITECTURE.md) for the request paths and trust boundaries.

## Scope

This repository is intended for code review and local verification, not as a
turnkey deployment. Runtime credentials, generated agent workspaces, live-state
services, and deployment helpers remain private. The public tests cover the
connector's identity bridge, workflow controls, query filtering, deduplication,
and date handling. CI builds and tests the TypeScript plugin, runs the Python
demo on Linux and Windows,
audits runtime dependencies, and scans Git history for secrets.

## License

[MIT](LICENSE)
