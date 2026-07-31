import { appendFileSync, copyFileSync, existsSync, mkdirSync, readFileSync, readdirSync, renameSync, statSync, unlinkSync, writeFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { Type } from "typebox";
import { defineToolPlugin } from "openclaw/plugin-sdk/tool-plugin";
import { jsonResult, type AnyAgentTool, type OpenClawPluginToolContext } from "openclaw/plugin-sdk/core";

import {
  PROJECT_ROOT, PRIVATE_ROOT, ROSTER_PATH, PRIVATE_USER_DIRECTORY_PATH, WORKFLOWS_ROOT, WORKFLOW_RUNS_ROOT,
  CALENDAR_WORKFLOW_STATE_ROOT, MAIL_WORKFLOW_STATE_ROOT, TOKEN_DIR, M365_CLIENT_CONFIG_PATH, WORKFLOW_EXECUTOR_PATH,
  M365_HEALTH_HELPER_PATH, AUDIT_PATH, DELEGATED_WORKFLOW_RECEIPT_INDEX_PATH, CALENDAR_QUERY_AUDIT_PATH,
  EMAIL_QUERY_AUDIT_PATH, TEAM_CALENDAR_AUDIT_PATH, BACKUP_ROOT, EMAIL_ALERT_POLICY_PATH, LAUNCH_AGENTS_DIR,
  PREFERENCES_ROOT, OPENCLAW_CONFIG_PATH, STATE_REDACTED_USERS_ROOT, PYTHON_BIN, PER_MEMBER_AGENT_REFS,
  APPROVED_SAFE_DISPLAY_NAMES_BY_REF, EMAIL_ALERT_MIN_CADENCE_SECONDS, DEFAULT_EMAIL_ALERT_ACTIVE_QUOTA,
  MIN_EMAIL_ALERT_ACTIVE_QUOTA, MAX_EMAIL_ALERT_ACTIVE_QUOTA,
} from "./config.js";
import { workflowIdSchema, callerSchema, specSchema, normalizedCreateParams } from "./schemas.js";
import {
  allowedMemberCalendarOperations, deltaCalendarOperations, alertCalendarOperations,
  allowedMemberMailOperations, deltaMailOperations, alertMailOperations,
} from "./operations.js";
import {
  normalizeUserRef, tokenFileNameFor, tokenFileExistsByStat, redact, okResult, errorResult, safeFileNamePart,
  readJsonFile, writeJsonAtomic, ensureDir, writePrivateJsonAtomic, isoNow,
} from "./util.js";
import {
  ensurePreferenceRecord, forgetMyPreferenceForCaller, pendingPreferencePathFor,
  preferencePathFor, setMyPreferenceForCaller, showMyPreferencesForCaller,
  suggestMyPreferenceForCaller,
} from "./preferences.js";
import {
  hashCandidates, loadRoster, requireAdmin, requireSelfCapable, resolveCaller, sha256Hex,
} from "./identity.js";
import {
  dryM365EmailQueryForCaller, emailQuerySelectFields, emailQuestionTypeFromArgs,
  emailWindowFromArgs, m365EmailQueryForCaller, normalizeSearchText,
} from "./email.js";
import {
  activeEmailAlertCountsByOwner, adminEmailAlertQuotaGet, adminEmailAlertQuotaSet,
  adminWorkflowChangeForMember, adminWorkflowCreateForMember, adminWorkflowInspectForMember,
  adminWorkflowInventory, adminWorkflowTerminateForMember, adminWorkflowUpdateForMember,
  applyEmailAlertCadenceFloor, approvePendingAction, changeAutomation,
  changeMyWorkflowForCaller, createWorkflowSpec, createWorkflowSpecForCaller,
  detectSelfWorkflowTargetConflict, emailAlertDedupKey, findActiveEmailAlertDuplicateForOwner,
  inspectAutomation, listAllAutomations, listMyWorkflowsForCaller,
  m365HealthAllConnectedAccounts, myEmailAlertQuotaForCaller, myWorkflowStatusForCaller,
  readEmailAlertPolicy, resolveDelegatedMemberTarget, runMyWorkflowForCaller,
  terminateAutomation, terminateMyWorkflowForCaller, terminateWorkflowForOwner,
} from "./workflows.js";
import {
  activePreferenceValue, addCalendarDays, calendarDescriptorsFromGraphData,
  calendarGroupIdsFromGraphData, calendarQuerySelectFields, calendarWindowFromArgs,
  calendarWindowLabel, classifyCalendarOwnership, datedPartsInTimeZone,
  dryM365CalendarQueryForCaller, eventOverlapsWindow, graphCalendarGroupsListUrl,
  graphCalendarListUrl, graphCalendarsInGroupUrl, graphCalendarViewUrl,
  ianaFromGraphTimeZone, loadTokenWrapperForGraph, accessTokenFromLoad,
  m365CalendarQueryForCaller, m365TeamAvailabilityForCaller, normalizeTimezoneName,
  ownCalendarIdsFromGraphData, parseExplicitDateRange, previousMondayParts,
  questionTypeFromArgs, refreshAccessToken, safeCalendarText, safeUserDisplayName,
  teammateDisplayName, teamScheduleEventsInWindow, timezoneAwareArgs,
  userTimeZoneToIana, zonedDayStartInstant,
  type GraphFetchResult,
} from "./calendar.js";
import type {
  Role, RosterUser, WorkflowSpec, CalendarQueryWindow, CalendarQuestionType, CalendarMetadataEvent,
  TeamAvailabilityIntent, TeamScheduleTarget, EmailQueryWindow, EmailQuestionType, EmailMetadataMessage,
  EmailQueryResolvedWindow, EmailQueryFilters, EmailFolderScope, Caller, ToolResult,
} from "./types.js";

const toolsAdded = [
  "create_my_workflow",
  "list_my_workflows",
  "run_my_workflow",
  "pause_my_workflow",
  "resume_my_workflow",
  "terminate_my_workflow",
  "my_workflow_status",
  "my_email_alert_quota",
  "m365_calendar_query",
  "m365_email_query",
  "m365_team_availability",
  "show_my_preferences",
  "set_my_preference",
  "forget_my_preference",
  "suggest_my_preference",
  "create_workflow",
  "admin_workflow_create_for_member",
  "admin_workflow_update_for_member",
  "admin_workflow_pause_for_member",
  "admin_workflow_resume_for_member",
  "admin_workflow_terminate_for_member",
  "admin_workflow_inspect_for_member",
  "list_all_automations",
  "admin_workflow_inventory",
  "admin_email_alert_quota_get",
  "admin_email_alert_quota_set",
  "pause_automation",
  "resume_automation",
  "terminate_automation",
  "inspect_automation",
  "approve_pending_action",
  "m365_health_all_connected_accounts",
];

const callerParam = { caller_user_ref: callerSchema };
const workflowIdParam = { workflow_id: workflowIdSchema };
const optionalWorkflowIdParam = { workflow_id: Type.Optional(Type.String({ description: "Workflow identifier. For delegated admin follow-ups, omit this to use the latest persisted delegated workflow receipt; Slack message timestamps are rejected as workflow ids." })) };
const optionalCallerParam = { caller_user_ref: Type.Optional(Type.String({ description: "Deprecated for member tools. EMClaw resolves caller identity from trusted Slack/runtime context; conflicting values fail closed." })) };
const calendarQueryParams = {
  ...optionalCallerParam,
  question: Type.Optional(Type.String({ description: "Natural-language calendar basic-details question from the caller. Do not include user_ref." })),
  window: Type.Optional(Type.String({ description: "today, tomorrow, yesterday, this_afternoon, next_24_hours, last_7_days, this_week, last_week, or explicit range via startDateTime/endDateTime." })),
  startDateTime: Type.Optional(Type.String({ description: "Explicit ISO start for a basic calendar details range." })),
  endDateTime: Type.Optional(Type.String({ description: "Explicit ISO exclusive end for a basic calendar details range." })),
  dateText: Type.Optional(Type.String({ description: "Natural date phrase such as May 27, last week, this week, next Monday, previous work week, or week of May 25." })),
  question_type: Type.Optional(Type.String({ description: "day_brief, important_meetings, next_meeting, or schedule_summary." })),
  timezone: Type.Optional(Type.String({ description: "IANA timezone. Defaults to America/New_York." })),
};

const emailQueryParams = {
  ...optionalCallerParam,
  question: Type.Optional(Type.String({ description: "Natural-language email basic-details question from the caller. Do not include user_ref." })),
  question_type: Type.Optional(Type.String({ description: "morning_brief, attention_digest, recent_unread, today_digest, or since_yesterday." })),
  window: Type.Optional(Type.String({ description: "today, yesterday, since_yesterday, last_24_hours, last_48_hours, last_7_days, last_30_days, this_week, last_week, or explicit range via startDateTime/endDateTime." })),
  startDateTime: Type.Optional(Type.String({ description: "Explicit ISO start for a basic-details mail date range." })),
  endDateTime: Type.Optional(Type.String({ description: "Explicit ISO exclusive end for a basic-details mail date range." })),
  timezone: Type.Optional(Type.String({ description: "Timezone for natural date parsing. Defaults to America/New_York." })),
  dateText: Type.Optional(Type.String({ description: "Natural date phrase such as May 27, last week, last Monday-Friday, previous work week, or week of May 25." })),
  senderOrDomain: Type.Optional(Type.String({ description: "Safe local filter against selected sender display/domain details only. Raw email addresses are never returned." })),
  subjectContains: Type.Optional(Type.String({ description: "Safe local filter against selected subject only; no body search." })),
  keyword: Type.Optional(Type.String({ description: "Safe local keyword filter against selected subject and safe sender/domain labels only." })),
  max_results: Type.Optional(Type.Number({ description: "Maximum messages to inspect from selected metadata. Default 10, hard cap 50." })),
  folderScope: Type.Optional(Type.String({ description: "inbox or mailbox_basic. mailbox_basic uses selected basic message fields only." })),
  folder: Type.Optional(Type.String({ description: "Legacy alias for folderScope; inbox by default." })),
};

const teamAvailabilityParams = {
  ...optionalCallerParam,
  question: Type.Optional(Type.String({ description: "Natural-language team calendar availability or ownership question. Do not include user_ref." })),
  intent: Type.Optional(Type.String({ description: "team_free_busy, find_mutual_time, only_my_events, shared_events, or classify_event." })),
  window: Type.Optional(Type.String({ description: "today, tomorrow, yesterday, this_afternoon, next_24_hours, last_7_days, this_week, last_week, or explicit range via startDateTime/endDateTime." })),
  startDateTime: Type.Optional(Type.String({ description: "Explicit ISO start for free/busy or basic calendar details range." })),
  endDateTime: Type.Optional(Type.String({ description: "Explicit ISO exclusive end for free/busy or basic calendar details range." })),
  dateText: Type.Optional(Type.String({ description: "Natural date phrase such as May 27, last week, previous work week, or week of May 25." })),
  timezone: Type.Optional(Type.String({ description: "IANA timezone. Defaults to America/New_York." })),
  team_refs: Type.Optional(Type.Array(Type.String({ description: "Approved redacted roster user refs only. Omit for approved team roster." }))),
  dry_run: Type.Optional(Type.Boolean({ description: "Use local synthetic events only for dry validation." })),
  synthetic_events: Type.Optional(Type.Array(Type.Any({ description: "Synthetic basic-calendar-detail events for dry validation. Never include bodies, attendee emails, or links." }))),
};

const preferenceToolParams = {
  ...optionalCallerParam,
  key: Type.Optional(Type.String({ description: "Preference key, such as tone_preference, response_format_preference, display_name_preference, timezone_preference (an IANA zone like Asia/Singapore or a UTC/GMT offset like GMT+8 used to interpret calendar/email dates and times), email_digest_preferences, email_alert_preferences, calendar_alert_preferences, priority_senders_or_domains, subject_keyword_rules, or calendar_keyword_rules." })),
  value: Type.Optional(Type.String({ description: "Safe preference value or label. Do not include email bodies, raw email addresses, tokens, or Microsoft 365 content." })),
  source: Type.Optional(Type.String({ description: "Preference source: explicit or confirmed_inferred. Defaults to explicit for set_my_preference." })),
};

const adminInventoryParams = {
  ...optionalCallerParam,
  detail: Type.Optional(Type.String({ description: "Admin inventory mode: counts, active_records, terminated_records, or roster_roles." })),
  search: Type.Optional(Type.String({ description: "Optional safe search text for active_records or terminated_records, such as workflow id, name, operation, or subject filter." })),
  workflow_id: Type.Optional(Type.String({ description: "Optional workflow id filter for active_records or terminated_records." })),
  name: Type.Optional(Type.String({ description: "Optional workflow name filter for active_records or terminated_records." })),
  query: Type.Optional(Type.String({ description: "Optional query filter for active_records or terminated_records." })),
};

const adminEmailAlertQuotaParams = {
  ...optionalCallerParam,
  max_active_email_alerts_per_user: Type.Optional(Type.Number({ description: "Company-wide maximum active email alert workflows per user. Whole number from 1 to 5." })),
  max_active: Type.Optional(Type.Number({ description: "Alias for max_active_email_alerts_per_user." })),
  max: Type.Optional(Type.Number({ description: "Alias for max_active_email_alerts_per_user." })),
  value: Type.Optional(Type.Number({ description: "Alias for max_active_email_alerts_per_user." })),
};

const delegatedTargetParams = {
  ...optionalCallerParam,
  target_member: Type.Optional(Type.String({ description: "Safe roster display or alias, such as Avery or Cameron. Ambiguous values fail closed." })),
  target_member_ref: Type.Optional(Type.String({ description: "Safe target member reference; exact user_ref values require confirm_target_user_ref." })),
  target_user_ref: Type.Optional(Type.String({ description: "Safe target member reference; exact user_ref values require confirm_target_user_ref." })),
  target_agent_id: Type.Optional(Type.String({ description: "Target member agent id from the approved roster." })),
  confirm_target_user_ref: Type.Optional(Type.String({ description: "Required when an exact redacted user_ref is supplied directly." })),
};

const delegatedWorkflowParams = { ...delegatedTargetParams, ...optionalWorkflowIdParam };
const delegatedCreateParams = { ...delegatedTargetParams, ...normalizedCreateParams };
const delegatedUpdateParams = {
  ...delegatedWorkflowParams,
  patch_json: Type.Optional(Type.String({ description: "JSON object containing only allowed delegated workflow edits: schedule, criteria, enabled, delivery, labels, description, name, timezone, or trigger." })),
  schedule: Type.Optional(Type.Any()),
  criteria: Type.Optional(Type.Any()),
  enabled: Type.Optional(Type.Boolean()),
  delivery: Type.Optional(Type.Any()),
  labels: Type.Optional(Type.Any()),
  description: Type.Optional(Type.String()),
  name: Type.Optional(Type.String()),
  timezone: Type.Optional(Type.String()),
  trigger: Type.Optional(Type.Any()),
};

type TrustedContextSummary = {
  agentId?: string;
  sessionKey?: string;
  requesterSenderId?: string;
  deliveryTo?: string;
};

function sessionPeerCandidate(sessionKey: unknown): string | undefined {
  if (typeof sessionKey !== "string") return undefined;
  const parts = sessionKey.split(":");
  const directIndex = parts.indexOf("direct");
  if (directIndex >= 0 && typeof parts[directIndex + 1] === "string" && parts[directIndex + 1].trim()) return parts[directIndex + 1];
  return undefined;
}

function contextSummary(context: OpenClawPluginToolContext): TrustedContextSummary {
  const delivery = context.deliveryContext as { to?: unknown } | undefined;
  return {
    agentId: typeof context.agentId === "string" ? context.agentId : undefined,
    sessionKey: typeof context.sessionKey === "string" ? context.sessionKey : undefined,
    requesterSenderId: typeof context.requesterSenderId === "string" ? context.requesterSenderId : undefined,
    deliveryTo: typeof delivery?.to === "string" ? delivery.to : undefined,
  };
}

function resolveRosterUserFromContext(context: OpenClawPluginToolContext): Caller {
  const summary = contextSummary(context);
  const roster = loadRoster();
  const hashSet = new Set<string>();
  for (const candidate of [summary.requesterSenderId, summary.deliveryTo, sessionPeerCandidate(summary.sessionKey)]) {
    for (const hash of hashCandidates(candidate)) hashSet.add(hash);
  }
  const matches = roster.filter((user) => {
    if (user.status !== "active") return false;
    if (user.slack_user_sha256 && hashSet.has(user.slack_user_sha256)) return true;
    if (user.slack_user_hash_short && Array.from(hashSet).some((hash) => hash.startsWith(user.slack_user_hash_short ?? ""))) return true;
    return false;
  });
  const uniqueRefs = Array.from(new Set(matches.map((user) => user.user_ref)));
  if (uniqueRefs.length === 1) {
    const user = matches.find((entry) => entry.user_ref === uniqueRefs[0])!;
    return { user_ref: user.user_ref, role: user.role, status: user.status };
  }
  if (uniqueRefs.length > 1) throw new Error("caller_identity_ambiguous");
  const trustedAgentId = summary.agentId || "";
  if (trustedAgentId && PER_MEMBER_AGENT_REFS[trustedAgentId]) {
    const mappedRef = PER_MEMBER_AGENT_REFS[trustedAgentId];
    const mapped = roster.find((entry) => entry.user_ref === mappedRef && entry.status === "active");
    if (mapped) return { user_ref: mapped.user_ref, role: mapped.role, status: mapped.status };
  }
  if (summary.agentId === "emclaw-admin-v2") {
    const admin = roster.find((entry) => entry.user_ref === "admin" && entry.status === "active");
    if (admin) return { user_ref: admin.user_ref, role: admin.role, status: admin.status };
  }
  throw new Error("caller_identity_unresolved");
}

function resolveTrustedAdminCaller(context: OpenClawPluginToolContext): Caller {
  const caller = resolveRosterUserFromContext(context);
  requireAdmin(caller);
  return caller;
}

function resolveTrustedMemberCaller(params: unknown, context: OpenClawPluginToolContext): Caller {
  const caller = resolveRosterUserFromContext(context);
  requireSelfCapable(caller);
  const supplied = params && typeof params === "object" ? (params as { caller_user_ref?: unknown }).caller_user_ref : undefined;
  if (typeof supplied === "string" && supplied.trim() && supplied.trim() !== caller.user_ref) {
    throw new Error("caller_user_ref_conflicts_with_trusted_identity");
  }
  return caller;
}

function makeTrustedMemberTool(params: {
  name: string;
  label: string;
  description: string;
  parameters: ReturnType<typeof Type.Object>;
  run: (args: Record<string, unknown>, caller: Caller) => ToolResult | Promise<ToolResult>;
}): { name: string; label: string; description: string; parameters: ReturnType<typeof Type.Object>; optional: true; factory: (context: { toolContext: OpenClawPluginToolContext }) => AnyAgentTool } {
  return {
    name: params.name,
    label: params.label,
    description: params.description,
    parameters: params.parameters,
    optional: true,
    factory: ({ toolContext }) => ({
      name: params.name,
      label: params.label,
      description: params.description,
      parameters: params.parameters,
      execute: async (_toolCallId: string, args: unknown, signal?: AbortSignal) => {
        signal?.throwIfAborted();
        try {
          const record = args && typeof args === "object" ? args as Record<string, unknown> : {};
          const caller = resolveTrustedMemberCaller(record, toolContext);
          return jsonResult(await params.run(record, caller));
        } catch (err) {
          return jsonResult(errorResult("trusted_caller_resolution_failed", err instanceof Error ? err.message : String(err)));
        }
      },
    }),
  };
}


function makeTrustedAdminTool(params: {
  name: string;
  label: string;
  description: string;
  parameters: ReturnType<typeof Type.Object>;
  run: (args: Record<string, unknown>, caller: Caller) => ToolResult | Promise<ToolResult>;
}): { name: string; label: string; description: string; parameters: ReturnType<typeof Type.Object>; optional: true; factory: (context: { toolContext: OpenClawPluginToolContext }) => AnyAgentTool } {
  return {
    name: params.name,
    label: params.label,
    description: params.description,
    parameters: params.parameters,
    optional: true,
    factory: ({ toolContext }) => ({
      name: params.name,
      label: params.label,
      description: params.description,
      parameters: params.parameters,
      execute: async (_toolCallId: string, args: unknown, signal?: AbortSignal) => {
        signal?.throwIfAborted();
        try {
          const record = args && typeof args === "object" ? args as Record<string, unknown> : {};
          return jsonResult(await params.run(record, resolveTrustedAdminCaller(toolContext)));
        } catch (err) {
          return jsonResult(errorResult(params.name + "_failed", err instanceof Error ? err.message : String(err)));
        }
      },
    }),
  };
}

export default defineToolPlugin({
  id: "emclaw-m365-connector",
  name: "EMClaw Microsoft 365 Connector",
  description: "Deterministic Microsoft 365 connector and IOC workflow wrapper tools for EMClaw IOC 1.0.",
  tools: (tool) => [
    tool({
      name: "m365_status",
      label: "Microsoft 365 Status",
      description: "Return deterministic redacted Microsoft 365 connector status by private token-file presence. No Graph, OAuth, token, email, calendar, OneDrive, or SharePoint content access.",
      parameters: Type.Object({ user_ref: Type.Optional(Type.String({ description: "Admin may inspect a redacted roster user_ref. Members should omit this; EMClaw resolves the caller from trusted runtime context." })) }),
      optional: true,
      async execute(params, _config, context) {
        context.signal?.throwIfAborted();
        const supplied = (params as { user_ref?: unknown } | undefined)?.user_ref;
        let trustedCaller: Caller | null = null;
        try { trustedCaller = resolveRosterUserFromContext(context as unknown as OpenClawPluginToolContext); } catch { trustedCaller = null; }
        const suppliedUserRef = typeof supplied === "string" && supplied.trim() ? normalizeUserRef(supplied) : "";
        const userRef = suppliedUserRef || trustedCaller?.user_ref || "admin";
        if (trustedCaller && trustedCaller.role !== "admin" && userRef !== trustedCaller.user_ref) {
          return errorResult("caller_user_ref_conflicts_with_trusted_identity", "Microsoft 365 status is caller-bound for members; I can only check your own connected account.", { trusted_identity_resolved: true });
        }
        const connected = tokenFileExistsByStat(userRef);
        return okResult({
          handled: true,
          connector_id: "microsoft365",
          operation: "m365_status",
          user_ref: userRef,
          trusted_identity_resolved: Boolean(trustedCaller),
          caller_bound_status_check: trustedCaller?.role !== "admin",
          connected,
          state: connected ? "connected_token_private" : "not_connected",
          checked_by: "token_file_presence_stat_only",
          token_file_present: connected,
          private_token_cache_contents_read: false,
          graph_calls_performed: false,
          oauth_started: false,
          m365_content_read: false,
          live_effect_performed: false,
          runtime_feature_accepted: true,
          next_human_action: connected ? null : "Run the admin-assisted Microsoft 365 connection helper for this user_ref.",
        });
      },
    }),
    tool(makeTrustedMemberTool({
      name: "create_my_workflow",
      label: "Create My Workflow",
      description: "Create the current routed member's own single-user workflow after executor validation. Product-facing skills must pass explicit app_id, service, operation, scope, schedule, delivery, and content_access metadata. EMClaw resolves caller identity from trusted Slack/runtime context; the model must not supply or choose caller_user_ref.",
      parameters: Type.Object({ ...optionalCallerParam, ...normalizedCreateParams }),
      run: (args, caller) => createWorkflowSpecForCaller(args, caller, false),
    })),
    tool(makeTrustedMemberTool({
      name: "list_my_workflows",
      label: "List My Workflows",
      description: "List workflow metadata for the current routed member. EMClaw resolves caller identity from trusted Slack/runtime context.",
      parameters: Type.Object(optionalCallerParam),
      run: (_args, caller) => listMyWorkflowsForCaller(caller),
    })),
    tool(makeTrustedMemberTool({
      name: "run_my_workflow",
      label: "Dry Run My Workflow",
      description: "Run or preview the current routed member's own workflow through the private executor where implementation support exists. Own-inbox Slack email alerts use the accepted live alert path; unsupported sends/writes remain unavailable.",
      parameters: Type.Object({ ...optionalCallerParam, ...workflowIdParam }),
      run: (args, caller) => runMyWorkflowForCaller(caller, args.workflow_id),
    })),
    tool(makeTrustedMemberTool({
      name: "pause_my_workflow",
      label: "Pause My Workflow",
      description: "Set enabled=false for the current routed member's own workflow spec with backup; scheduling applies only where an accepted scheduler is installed.",
      parameters: Type.Object({ ...optionalCallerParam, ...workflowIdParam }),
      run: (args, caller) => changeMyWorkflowForCaller(caller, args.workflow_id, false, "pause_my_workflow"),
    })),
    tool(makeTrustedMemberTool({
      name: "resume_my_workflow",
      label: "Resume My Workflow",
      description: "Set enabled=true for the current routed member's own workflow spec with backup; scheduling applies only where an accepted scheduler is installed.",
      parameters: Type.Object({ ...optionalCallerParam, ...workflowIdParam }),
      run: (args, caller) => changeMyWorkflowForCaller(caller, args.workflow_id, true, "resume_my_workflow"),
    })),
    tool(makeTrustedMemberTool({
      name: "terminate_my_workflow",
      label: "Terminate My Workflow",
      description: "Move the current routed member's own workflow spec to quarantine with backup. No hard delete.",
      parameters: Type.Object({ ...optionalCallerParam, ...workflowIdParam }),
      run: (args, caller) => terminateMyWorkflowForCaller(caller, args.workflow_id),
    })),
    tool(makeTrustedMemberTool({
      name: "my_workflow_status",
      label: "My Workflow Status",
      description: "Return workflow metadata and recent run status metadata for the current routed member's own workflow.",
      parameters: Type.Object({ ...optionalCallerParam, ...workflowIdParam }),
      run: (args, caller) => myWorkflowStatusForCaller(caller, args.workflow_id),
    })),
    tool(makeTrustedMemberTool({
      name: "my_email_alert_quota",
      label: "My Email Alert Quota",
      description: "Return the current routed member's own email alert quota: the company-wide maximum active email alerts per user, the caller's current active email alert count, remaining slots, and the cadence floor. EMClaw resolves caller identity from trusted Slack/runtime context. Reads only the policy file and the caller's own active workflow specs; never returns other users' counts, calls Graph, performs OAuth, reads tokens or email/calendar content, sends Slack, or mutates workflows.",
      parameters: Type.Object(optionalCallerParam),
      run: (_args, caller) => myEmailAlertQuotaForCaller(caller),
    })),
    tool(makeTrustedMemberTool({
      name: "m365_calendar_query",
      label: "Microsoft 365 Calendar Query",
      description: "Answer the current routed caller's own one-time Microsoft 365 calendar question using a fresh tool call every time, including repeated or identical questions in the same chat. Use basic calendar details only, exact date ranges, and timezone-aware display. Uses trusted caller identity; no event notes, descriptions, attachments, attendee emails, meeting links, calendar writes, workflows, Slack sends, or token contents.",
      parameters: Type.Object(calendarQueryParams),
      run: (args, caller) => m365CalendarQueryForCaller(args, caller),
    })),
    tool(makeTrustedMemberTool({
      name: "m365_email_query",
      label: "Microsoft 365 Email Metadata Query",
      description: "Answer the current routed caller's own one-time Microsoft 365 email question using a fresh tool call every time, including repeated or identical questions in the same chat. Use basic email details only, exact date ranges, mailbox-basic or inbox scope, and safe local matching. Uses trusted caller identity; no bodies, body previews, uniqueBody, headers, MIME, attachments, raw email addresses, mailbox writes, workflows, Slack sends, or token contents.",
      parameters: Type.Object(emailQueryParams),
      run: (args, caller) => m365EmailQueryForCaller(args, caller),
    })),
    tool(makeTrustedMemberTool({
      name: "m365_team_availability",
      label: "Microsoft 365 Team Availability",
      description: "Answer team calendar availability and event ownership questions using approved roster refs and basic calendar details only. Defaults to free/busy only for members; never reads event notes, attachments, attendee emails, meeting links, or calendar writes.",
      parameters: Type.Object(teamAvailabilityParams),
      run: (args, caller) => m365TeamAvailabilityForCaller(args, caller),
    })),
    tool(makeTrustedMemberTool({
      name: "show_my_preferences",
      label: "Show My Preferences",
      description: "Show the current routed member's own structured EMClaw secretary preferences. Uses trusted caller identity; never reads Microsoft 365 content or token contents.",
      parameters: Type.Object(optionalCallerParam),
      run: (_args, caller) => showMyPreferencesForCaller(caller),
    })),
    tool(makeTrustedMemberTool({
      name: "set_my_preference",
      label: "Set My Preference",
      description: "Set an explicit or confirmed-inferred secretary preference for the current routed member only. Do not store raw Slack IDs, emails, tokens, or Microsoft 365 content.",
      parameters: Type.Object(preferenceToolParams),
      run: (args, caller) => setMyPreferenceForCaller(args, caller),
    })),
    tool(makeTrustedMemberTool({
      name: "forget_my_preference",
      label: "Forget My Preference",
      description: "Forget or mark forgotten one of the current routed member's own structured secretary preferences.",
      parameters: Type.Object(preferenceToolParams),
      run: (args, caller) => forgetMyPreferenceForCaller(args, caller),
    })),
    tool(makeTrustedMemberTool({
      name: "suggest_my_preference",
      label: "Suggest My Preference",
      description: "Record a pending inferred preference suggestion for the current routed member. Pending suggestions are not active until the member confirms them.",
      parameters: Type.Object(preferenceToolParams),
      run: (args, caller) => suggestMyPreferenceForCaller(args, caller),
    })),

    tool(makeTrustedAdminTool({
      name: "admin_workflow_create_for_member",
      label: "Admin Create Member Workflow",
      description: "Admin-only delegated workflow creation for a target member. The workflow is member-owned, runs as the target member, uses the target member's connected app account, and exposes only redacted config/status/audit to admin.",
      parameters: Type.Object(delegatedCreateParams),
      run: (args, caller) => adminWorkflowCreateForMember(caller, args),
    })),
    tool(makeTrustedAdminTool({
      name: "admin_workflow_update_for_member",
      label: "Admin Update Member Workflow",
      description: "Admin-only delegated workflow update for allowed config fields: schedule, criteria, enabled state, delivery mode, labels, description, name, timezone, and trigger. Does not expose private run payloads or M365 content.",
      parameters: Type.Object(delegatedUpdateParams),
      run: (args, caller) => adminWorkflowUpdateForMember(caller, args),
    })),
    tool(makeTrustedAdminTool({
      name: "admin_workflow_pause_for_member",
      label: "Admin Pause Member Workflow",
      description: "Admin-only pause for a target member-owned workflow. Preserves owner/runs_as target member semantics and records redacted audit.",
      parameters: Type.Object(delegatedWorkflowParams),
      run: (args, caller) => adminWorkflowChangeForMember(caller, args, false, "admin_workflow_pause_for_member"),
    })),
    tool(makeTrustedAdminTool({
      name: "admin_workflow_resume_for_member",
      label: "Admin Resume Member Workflow",
      description: "Admin-only resume for a target member-owned workflow when the target member connection is ready. Preserves owner/runs_as target member semantics.",
      parameters: Type.Object(delegatedWorkflowParams),
      run: (args, caller) => adminWorkflowChangeForMember(caller, args, true, "admin_workflow_resume_for_member"),
    })),
    tool(makeTrustedAdminTool({
      name: "admin_workflow_terminate_for_member",
      label: "Admin Terminate Member Workflow",
      description: "Admin-only quarantine termination for a target member-owned workflow. No hard delete; member namespace quarantine is used.",
      parameters: Type.Object(delegatedWorkflowParams),
      run: (args, caller) => adminWorkflowTerminateForMember(caller, args),
    })),
    tool(makeTrustedAdminTool({
      name: "admin_workflow_inspect_for_member",
      label: "Admin Inspect Member Workflow",
      description: "Admin-only redacted inspection of target member-owned workflow config/status. Does not return private M365 content, raw payloads, tokens, raw Slack IDs, or raw emails.",
      parameters: Type.Object(delegatedWorkflowParams),
      run: (args, caller) => adminWorkflowInspectForMember(caller, args),
    })),
    tool(makeTrustedAdminTool({
      name: "admin_email_alert_quota_get",
      label: "Get Email Alert Quota",
      description: "Admin-only read of the company-wide per-user active email alert workflow maximum and current active counts. Does not call Graph, read tokens, read email content, or mutate workflows.",
      parameters: Type.Object(optionalCallerParam),
      run: (_args, caller) => adminEmailAlertQuotaGet(caller),
    })),
    tool(makeTrustedAdminTool({
      name: "admin_email_alert_quota_set",
      label: "Set Email Alert Quota",
      description: "Admin-only update of the company-wide per-user active email alert workflow maximum. Bounds are 1 through 5. Lowering the maximum does not terminate existing workflows.",
      parameters: Type.Object(adminEmailAlertQuotaParams),
      run: (args, caller) => adminEmailAlertQuotaSet(caller, args),
    })),
    tool({
      name: "create_workflow",
      label: "Create Workflow",
      description: "Admin-only workflow creation wrapper. Cross-user workflow validation requires roster role=admin.",
      parameters: Type.Object({ ...callerParam, spec_json: specSchema }),
      optional: true,
      async execute(params, _config, context) { context.signal?.throwIfAborted(); return createWorkflowSpec((params as { spec_json?: unknown }).spec_json, (params as { caller_user_ref?: unknown }).caller_user_ref, true); },
    }),
    tool({
      name: "list_all_automations",
      label: "List All Automations",
      description: "Admin-only list of workflow spec metadata across owners.",
      parameters: Type.Object(callerParam),
      optional: true,
      async execute(params, _config, context) { context.signal?.throwIfAborted(); return listAllAutomations((params as { caller_user_ref?: unknown }).caller_user_ref); },
    }),
    tool({
      name: "admin_workflow_inventory",
      label: "Admin Workflow Inventory",
      description: "Admin-only deterministic local workflow inventory and redacted operational metadata. Use detail=counts for aggregate counts, detail=active_records to find active workflows by workflow id, operation, name, or subject filter, detail=terminated_records for quarantined/terminated workflow metadata, and detail=roster_roles for redacted roster/role/route metadata. Caller identity is resolved from trusted OpenClaw admin route/runtime context, not model-supplied user_ref. Does not call Graph, read tokens, read M365 content, mutate workflows, or use m365_status.",
      parameters: Type.Object(adminInventoryParams),
      optional: true,
      factory: ({ toolContext }) => ({
        name: "admin_workflow_inventory",
        label: "Admin Workflow Inventory",
        description: "Admin-only deterministic local workflow inventory and redacted operational metadata. Use detail=counts for aggregate counts, detail=active_records to find active workflows by workflow id, operation, name, or subject filter, detail=terminated_records for quarantined/terminated workflow metadata, and detail=roster_roles for redacted roster/role/route metadata. Caller identity is resolved from trusted OpenClaw admin route/runtime context, not model-supplied user_ref. Does not call Graph, read tokens, read M365 content, mutate workflows, or use m365_status.",
        parameters: Type.Object(adminInventoryParams),
        execute: async (_toolCallId: string, _args: unknown, signal?: AbortSignal) => {
          signal?.throwIfAborted();
          try {
            const record = _args && typeof _args === "object" ? _args as Record<string, unknown> : {};
            return jsonResult(adminWorkflowInventory(resolveTrustedAdminCaller(toolContext), record));
          } catch (err) {
            return jsonResult(errorResult("admin_workflow_inventory_failed", err instanceof Error ? err.message : String(err)));
          }
        },
      }),
    }),
    tool({
      name: "pause_automation",
      label: "Pause Automation",
      description: "Admin-only pause for any workflow spec with backup.",
      parameters: Type.Object({ ...callerParam, ...workflowIdParam }),
      optional: true,
      async execute(params, _config, context) { context.signal?.throwIfAborted(); return changeAutomation((params as { caller_user_ref?: unknown }).caller_user_ref, (params as { workflow_id?: unknown }).workflow_id, false, "pause_automation"); },
    }),
    tool({
      name: "resume_automation",
      label: "Resume Automation",
      description: "Admin-only resume for any workflow spec with backup.",
      parameters: Type.Object({ ...callerParam, ...workflowIdParam }),
      optional: true,
      async execute(params, _config, context) { context.signal?.throwIfAborted(); return changeAutomation((params as { caller_user_ref?: unknown }).caller_user_ref, (params as { workflow_id?: unknown }).workflow_id, true, "resume_automation"); },
    }),
    tool({
      name: "terminate_automation",
      label: "Terminate Automation",
      description: "Admin-only terminate for any workflow spec. Moves spec to quarantine with backup; no hard delete.",
      parameters: Type.Object({ ...callerParam, ...workflowIdParam }),
      optional: true,
      async execute(params, _config, context) { context.signal?.throwIfAborted(); return terminateAutomation((params as { caller_user_ref?: unknown }).caller_user_ref, (params as { workflow_id?: unknown }).workflow_id); },
    }),
    tool({
      name: "inspect_automation",
      label: "Inspect Automation",
      description: "Admin-only inspect of workflow spec metadata. Redacts sensitive-looking values.",
      parameters: Type.Object({ ...callerParam, ...workflowIdParam }),
      optional: true,
      async execute(params, _config, context) { context.signal?.throwIfAborted(); return inspectAutomation((params as { caller_user_ref?: unknown }).caller_user_ref, (params as { workflow_id?: unknown }).workflow_id); },
    }),
    tool({
      name: "approve_pending_action",
      label: "Approve Pending Action",
      description: "Admin-only placeholder for future approval flow. Returns not implemented until pending-action files exist.",
      parameters: Type.Object({ ...callerParam, run_id: Type.String(), approval_token: Type.String(), decision: Type.String() }),
      optional: true,
      async execute(params, _config, context) {
        context.signal?.throwIfAborted();
        const p = params as { caller_user_ref?: unknown; run_id?: unknown; approval_token?: unknown; decision?: unknown };
        return approvePendingAction(p.caller_user_ref, p.run_id, p.approval_token, p.decision);
      },
    }),
    tool({
      name: "m365_health_all_connected_accounts",
      label: "Microsoft 365 Health: All Connected Accounts",
      description: "Admin-only narrow Graph /me health check for every connected Microsoft 365 account. Calls only /v1.0/me and returns redacted hashes/status metadata.",
      parameters: Type.Object(callerParam),
      optional: true,
      async execute(params, _config, context) {
        context.signal?.throwIfAborted();
        return m365HealthAllConnectedAccounts((params as { caller_user_ref?: unknown }).caller_user_ref);
      },
    }),
  ],
});

export { toolsAdded };
export const __test = { createWorkflowSpecForCaller, terminateWorkflowForOwner, detectSelfWorkflowTargetConflict, adminWorkflowCreateForMember, adminWorkflowUpdateForMember, adminWorkflowChangeForMember, adminWorkflowTerminateForMember, adminWorkflowInspectForMember, adminEmailAlertQuotaGet, adminEmailAlertQuotaSet, myEmailAlertQuotaForCaller, readEmailAlertPolicy, activeEmailAlertCountsByOwner, emailAlertDedupKey, findActiveEmailAlertDuplicateForOwner, applyEmailAlertCadenceFloor, resolveDelegatedMemberTarget, resolveRosterUserFromContext, resolveTrustedMemberCaller, hashCandidates, adminWorkflowInventory, showMyPreferencesForCaller, setMyPreferenceForCaller, forgetMyPreferenceForCaller, suggestMyPreferenceForCaller, preferencePathFor, pendingPreferencePathFor, dryM365CalendarQueryForCaller, dryM365EmailQueryForCaller, m365TeamAvailabilityForCaller, safeUserDisplayName, teammateDisplayName, classifyCalendarOwnership, calendarWindowFromArgs, calendarWindowLabel, parseExplicitDateRange, ianaFromGraphTimeZone, userTimeZoneToIana, datedPartsInTimeZone, questionTypeFromArgs, calendarQuerySelectFields, graphCalendarViewUrl, graphCalendarListUrl, graphCalendarGroupsListUrl, graphCalendarsInGroupUrl, ownCalendarIdsFromGraphData, eventOverlapsWindow, teamScheduleEventsInWindow, calendarDescriptorsFromGraphData, calendarGroupIdsFromGraphData, emailWindowFromArgs, emailQuestionTypeFromArgs, emailQuerySelectFields };
