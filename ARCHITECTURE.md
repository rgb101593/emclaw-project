# Architecture

EMClaw has two execution paths: interactive requests from Slack and scheduled
workflow runs. Both paths keep identity and authorization outside the language
model.

## Interactive requests

1. OpenClaw receives a Slack direct message and applies the channel allow-list.
2. Runtime routing selects the member's isolated agent.
3. A deployment hook handles connection and status requests that do not need
   model reasoning. That deployment-specific hook is not part of this snapshot.
4. The local model may answer directly or request a tool.
5. Before a tool runs, the connector resolves the sender from trusted runtime
   context and checks the requested operation against the caller's role.
6. Microsoft Graph is called with the caller's delegated token and restricted
   metadata fields.
7. Before delivery, output guards reject credential-shaped text and claims that
   lack deterministic evidence.
8. A redacted audit event is written to the private runtime tree.

Prompt text and caller-supplied tool arguments are not accepted as identity.
The roster stores one-way sender hashes, and unknown senders fail closed.

## Scheduled workflows

Scheduled email alerts run without an active conversation. The public connector
owns workflow validation, authorization, quotas, lifecycle state, and receipts.
A deployment-specific executor in the private runtime tree:

- loads the workflow and its private delta watermark;
- queries Microsoft Graph with the workflow owner's delegated token;
- filters and de-duplicates matching metadata;
- sends the alert to the approved Slack destination; and
- advances the watermark only after the complete cycle succeeds.

A partial failure does not advance the watermark. Termination moves workflow
state into quarantine and creates rollback instructions rather than deleting it
immediately. The private executor and live state are not included in this
sanitized snapshot.

## Public and private state

| Public repository | Private runtime tree |
|---|---|
| Connector source and tests | OAuth tokens |
| Redacted roster hashes | Names and email addresses |
| Security walkthrough | Audit logs and delta watermarks |
| CI and secret-scanning rules | Live workflow state and quarantine records |

`EMCLAW_ROOT` identifies the checkout. `EMCLAW_PRIVATE_ROOT` identifies the
separate private tree. The public snapshot does not include enough configuration
to connect to a real Slack or Microsoft 365 tenant.

## Connector modules

- `index.ts`: tool schemas, registration, and trusted-context adapters
- `identity.ts`: roster parsing, caller resolution, and role checks
- `calendar.ts`: calendar window parsing, Graph reads, and safe summaries
- `email.ts`: email metadata queries, filters, and safe summaries
- `preferences.ts`: structured per-user preferences
- `workflows.ts`: member and administrator workflow lifecycle operations
- `config.ts`, `schemas.ts`, `types.ts`, `util.ts`: shared boundaries

The local model is not a security boundary. The public connector enforces
identity checks, API field selection, authorization, and workflow state
transitions. The private deployment adds routing and outbound inspection.
