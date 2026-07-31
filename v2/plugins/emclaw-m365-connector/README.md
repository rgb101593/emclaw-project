# Microsoft 365 connector

This OpenClaw tool plugin provides EMClaw's identity bridge, metadata-only email
and calendar queries, per-user preferences, and workflow controls.

The public package uses a redacted roster and expects credentials, audit data,
workflow execution state, and deployment helpers under `EMCLAW_PRIVATE_ROOT`.
It is not a standalone Microsoft 365 application.

## Verify

```bash
npm ci
npm run typecheck
npm run build
npm test
npm audit --omit=dev
```

The tests cover trusted caller resolution, member/admin boundaries, workflow
quotas and deduplication, safe query shaping, timezone handling, and preference
isolation.
