// Deployment variables override the self-located public and private roots.
// Private tokens, audit data, and state remain outside the repository.
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";
import { homedir } from "node:os";

const moduleDir = dirname(fileURLToPath(import.meta.url));
// config.ts runs from src or dist; four parent directories reach the repository root.
const DERIVED_PROJECT_ROOT = resolve(moduleDir, "../../../..");

export const PROJECT_ROOT = process.env.EMCLAW_ROOT ?? DERIVED_PROJECT_ROOT;
export const PRIVATE_ROOT = process.env.EMCLAW_PRIVATE_ROOT ?? resolve(PROJECT_ROOT, "../emclaw-private");

export const ROSTER_PATH = `${PROJECT_ROOT}/v2/identity/roster.json`;
export const PRIVATE_USER_DIRECTORY_PATH = `${PRIVATE_ROOT}/v2/identity/private-user-directory.json`;
export const WORKFLOWS_ROOT = `${PROJECT_ROOT}/v2/workflows`;
export const WORKFLOW_RUNS_ROOT = `${PROJECT_ROOT}/v2/workflow-runs`;
export const CALENDAR_WORKFLOW_STATE_ROOT = `${PRIVATE_ROOT}/v2/workflow-state/m365-calendar`;
export const MAIL_WORKFLOW_STATE_ROOT = `${PRIVATE_ROOT}/v2/workflow-state/m365-mail`;
export const TOKEN_DIR = `${PRIVATE_ROOT}/v2/credentials/microsoft365`;
export const M365_CLIENT_CONFIG_PATH = `${PRIVATE_ROOT}/v2/config/microsoft365-client.json`;
export const WORKFLOW_EXECUTOR_PATH = `${PRIVATE_ROOT}/v2/services/workflow_executor.py`;
export const M365_HEALTH_HELPER_PATH = `${PRIVATE_ROOT}/v2/services/m365_health_all_connected_accounts.py`;
export const AUDIT_PATH = `${PRIVATE_ROOT}/v2/audit/workflows.jsonl`;
export const DELEGATED_WORKFLOW_RECEIPT_INDEX_PATH = `${PRIVATE_ROOT}/v2/workflow-receipts/admin-delegated-latest.json`;
export const CALENDAR_QUERY_AUDIT_PATH = `${PRIVATE_ROOT}/v2/audit/calendar-metadata-queries.jsonl`;
export const EMAIL_QUERY_AUDIT_PATH = `${PRIVATE_ROOT}/v2/audit/email-metadata-queries.jsonl`;
export const TEAM_CALENDAR_AUDIT_PATH = `${PRIVATE_ROOT}/v2/audit/team-calendar-availability.jsonl`;
export const BACKUP_ROOT = `${PRIVATE_ROOT}/v2/backups/workflow-tool-wrapper-runtime`;
export const EMAIL_ALERT_POLICY_PATH = `${PRIVATE_ROOT}/v2/policy/email-alert-quota.json`;
export const LAUNCH_AGENTS_DIR = process.env.EMCLAW_LAUNCH_AGENTS_DIR ?? `${homedir()}/Library/LaunchAgents`;
export const PREFERENCES_ROOT = `${PRIVATE_ROOT}/v2/preferences`;
export const OPENCLAW_CONFIG_PATH = process.env.OPENCLAW_CONFIG_PATH ?? `${homedir()}/.openclaw/openclaw.json`;
export const STATE_REDACTED_USERS_ROOT = `${PROJECT_ROOT}/state-redacted/users`;
export const PYTHON_BIN = process.env.PYTHON_BIN ?? "python3";

export const PER_MEMBER_AGENT_REFS: Record<string, string> = {
  "emclaw-member-8d4e38d9a1cf7e80-v2": "slack-user-ref-8d4e38d9a1cf7e80",
  "emclaw-member-db6b1d5cccad3cbf-v2": "slack-user-ref-db6b1d5cccad3cbf",
};
export const APPROVED_SAFE_DISPLAY_NAMES_BY_REF: Record<string, string> = {
  "slack-user-ref-8d4e38d9a1cf7e80": "Avery",
  "slack-user-ref-db6b1d5cccad3cbf": "Cameron",
};

export const EMAIL_ALERT_MIN_CADENCE_SECONDS = 180;
export const DEFAULT_EMAIL_ALERT_ACTIVE_QUOTA = 2;
export const MIN_EMAIL_ALERT_ACTIVE_QUOTA = 1;
export const MAX_EMAIL_ALERT_ACTIVE_QUOTA = 5;
