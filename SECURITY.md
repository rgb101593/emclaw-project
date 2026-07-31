# Security policy

## Reporting

Please report suspected vulnerabilities privately through GitHub's security
advisory interface rather than opening a public issue.

## Public snapshot boundary

This repository is a sanitized engineering snapshot. Live credentials, raw
provider identifiers, OAuth tokens, audit records, and workflow state belong in
the separate private runtime tree and must never be committed here.

The connector treats OpenClaw as a host-provided peer dependency. Its pinned
development copy exists only to compile and test the plugin. Release
verification therefore:

- builds and runs the complete test suite against that pinned host version;
- fails on high-severity vulnerabilities in shipped runtime dependencies; and
- scans the complete Git history with Gitleaks.

Upstream advisories in development-only dependencies are reviewed when the
OpenClaw pin is updated. They are not silently overridden across OpenClaw's
published shrinkwrap because doing so could invalidate the tested host runtime.
