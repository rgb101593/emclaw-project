import {
  appendFileSync, copyFileSync, existsSync, mkdirSync, readFileSync, readdirSync,
  renameSync, statSync, unlinkSync, writeFileSync,
} from "node:fs";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";

import {
  PROJECT_ROOT, PRIVATE_ROOT, ROSTER_PATH, PRIVATE_USER_DIRECTORY_PATH, WORKFLOWS_ROOT, WORKFLOW_RUNS_ROOT,
  CALENDAR_WORKFLOW_STATE_ROOT, MAIL_WORKFLOW_STATE_ROOT, TOKEN_DIR, M365_CLIENT_CONFIG_PATH, WORKFLOW_EXECUTOR_PATH,
  M365_HEALTH_HELPER_PATH, AUDIT_PATH, DELEGATED_WORKFLOW_RECEIPT_INDEX_PATH, CALENDAR_QUERY_AUDIT_PATH,
  EMAIL_QUERY_AUDIT_PATH, TEAM_CALENDAR_AUDIT_PATH, BACKUP_ROOT, EMAIL_ALERT_POLICY_PATH, LAUNCH_AGENTS_DIR,
  PREFERENCES_ROOT, OPENCLAW_CONFIG_PATH, STATE_REDACTED_USERS_ROOT, PYTHON_BIN, PER_MEMBER_AGENT_REFS,
  APPROVED_SAFE_DISPLAY_NAMES_BY_REF, EMAIL_ALERT_MIN_CADENCE_SECONDS, DEFAULT_EMAIL_ALERT_ACTIVE_QUOTA,
  MIN_EMAIL_ALERT_ACTIVE_QUOTA, MAX_EMAIL_ALERT_ACTIVE_QUOTA,
} from "./config.js";
import { normalizedCreateParams } from "./schemas.js";
import {
  allowedMemberCalendarOperations, deltaCalendarOperations, alertCalendarOperations,
  allowedMemberMailOperations, deltaMailOperations, alertMailOperations,
} from "./operations.js";
import {
  normalizeUserRef, tokenFileNameFor, tokenFileExistsByStat, redact, okResult, errorResult, safeFileNamePart,
  readJsonFile, writeJsonAtomic, ensureDir, writePrivateJsonAtomic, isoNow,
} from "./util.js";
import { normalizeSearchText } from "./email.js";
import {
  loadRoster, requireAdmin, requireSelfCapable, resolveCaller, sha256Hex,
} from "./identity.js";
import type { Caller, Role, RosterUser, ToolResult, WorkflowSpec } from "./types.js";

export function assertWorkflowId(value: unknown): string {
  if (typeof value !== "string" || !/^[A-Za-z0-9._-]{1,100}$/.test(value)) throw new Error("invalid_workflow_id");
  return value;
}

export function workflowDirForOwner(ownerUserRef: string): string {
  return `${WORKFLOWS_ROOT}/${safeFileNamePart(ownerUserRef)}`;
}

export function workflowPathForOwner(ownerUserRef: string, workflowId: string): string {
  return `${workflowDirForOwner(ownerUserRef)}/${assertWorkflowId(workflowId)}.json`;
}

export function xmlEscape(value: unknown): string {
  return String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

export function launchdLabelForWorkflow(workflowId: string, ownerUserRef = "admin"): string {
  const ownerPart = safeFileNamePart(ownerUserRef).replace(/[^A-Za-z0-9.-]/g, "-");
  return `com.emclaw.workflow.mail-delta-${ownerPart}-${assertWorkflowId(workflowId).replace(/[^A-Za-z0-9.-]/g, "-")}`;
}

export function launchdPlistPathForWorkflow(workflowId: string, ownerUserRef = "admin"): string {
  return `${LAUNCH_AGENTS_DIR}/${launchdLabelForWorkflow(workflowId, ownerUserRef)}.plist`;
}

export function isSelfOwnedMailSlackAlertSpec(spec: WorkflowSpec, caller: Caller): boolean {
  const delivery = objectValue(spec.delivery);
  const operation = stringValue(spec.operation);
  return (caller.role === "admin" || caller.role === "member")
    && stringValue(spec.owner_user_ref || caller.user_ref) === caller.user_ref
    && stringValue(spec.created_by_user_ref || caller.user_ref) === caller.user_ref
    && (spec.runs_as_user_ref === undefined || stringValue(spec.runs_as_user_ref) === caller.user_ref)
    && spec.admin_delegated !== true
    && (spec.service === "m365.mail" || spec.service === "mail")
    && alertMailOperations.has(operation)
    && delivery.slack_send === true
    && (delivery.mode === "slack_send" || delivery.mode === "slack_alert" || delivery.mode === undefined)
    && spec.mailbox_write !== true
    && spec.calendar_write !== true
    && spec.body_read !== true
    && spec.bodyPreview_read !== true
    && spec.uniqueBody_read !== true
    && spec.mime_read !== true
    && spec.headers_read !== true
    && spec.attachments_read !== true
    && spec.raw_email_addresses_exposed !== true;
}

export function launchdDomain(): string {
  const uid = spawnSync("/usr/bin/id", ["-u"], { encoding: "utf8" });
  const value = typeof uid.stdout === "string" ? uid.stdout.trim() : "501";
  return `gui/${value || "501"}`;
}

export function installMailDeltaLaunchAgent(workflowPath: string, spec: WorkflowSpec): ToolResult {
  const workflowId = assertWorkflowId(spec.workflow_id);
  const ownerUserRef = stringValue(spec.owner_user_ref) || "admin";
  const label = launchdLabelForWorkflow(workflowId, ownerUserRef);
  const plistPath = launchdPlistPathForWorkflow(workflowId, ownerUserRef);
  const interval = Math.max(60, Number((objectValue(spec.trigger).seconds || 60)) || 60);
  mkdirSync(LAUNCH_AGENTS_DIR, { recursive: true, mode: 0o755 });
  const stdout = `/private/tmp/${label}.out.log`;
  const stderr = `/private/tmp/${label}.err.log`;
  const plist = `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n<plist version="1.0">\n<dict>\n  <key>Label</key>\n  <string>${xmlEscape(label)}</string>\n  <key>ProgramArguments</key>\n  <array>\n    <string>/usr/bin/python3</string>\n    <string>${xmlEscape(WORKFLOW_EXECUTOR_PATH)}</string>\n    <string>--workflow-file</string>\n    <string>${xmlEscape(workflowPath)}</string>\n    <string>--trigger</string>\n    <string>scheduled</string>\n    <string>--execute-mail-delta-alert</string>\n  </array>\n  <key>StartInterval</key>\n  <integer>${interval}</integer>\n  <key>RunAtLoad</key>\n  <true/>\n  <key>StandardOutPath</key>\n  <string>${xmlEscape(stdout)}</string>\n  <key>StandardErrorPath</key>\n  <string>${xmlEscape(stderr)}</string>\n</dict>\n</plist>\n`;
  writeFileSync(plistPath, plist, { mode: 0o644 });
  const domain = launchdDomain();
  spawnSync("/bin/launchctl", ["bootout", domain, plistPath], { encoding: "utf8" });
  const bootstrap = spawnSync("/bin/launchctl", ["bootstrap", domain, plistPath], { encoding: "utf8" });
  const kickstart = spawnSync("/bin/launchctl", ["kickstart", "-k", `${domain}/${label}`], { encoding: "utf8" });
  return { launchd_label: label, launchd_plist_path: plistPath, start_interval_seconds: interval, bootstrap_returncode: bootstrap.status, bootstrap_stderr: redact(bootstrap.stderr || ""), kickstart_returncode: kickstart.status, kickstart_stderr: redact(kickstart.stderr || ""), launchd_installed: bootstrap.status === 0 || /service already loaded/i.test(String(bootstrap.stderr || "")) };
}

export function uninstallMailDeltaLaunchAgent(workflowId: string, ownerUserRef = "admin"): ToolResult {
  const label = launchdLabelForWorkflow(workflowId, ownerUserRef);
  const plistPath = launchdPlistPathForWorkflow(workflowId, ownerUserRef);
  const domain = launchdDomain();
  const bootout = spawnSync("/bin/launchctl", ["bootout", domain, plistPath], { encoding: "utf8" });
  let legacyBootout: ReturnType<typeof spawnSync> | null = null;
  let legacyRemoved = false;
  if (ownerUserRef === "admin") {
    const legacyLabel = `com.emclaw.workflow.admin-mail-delta-${assertWorkflowId(workflowId).replace(/[^A-Za-z0-9.-]/g, "-")}`;
    const legacyPlistPath = `${LAUNCH_AGENTS_DIR}/${legacyLabel}.plist`;
    legacyBootout = spawnSync("/bin/launchctl", ["bootout", domain, legacyPlistPath], { encoding: "utf8" });
    if (existsSync(legacyPlistPath)) {
      try { unlinkSync(legacyPlistPath); legacyRemoved = true; } catch { legacyRemoved = false; }
    }
  }
  let removed = false;
  if (existsSync(plistPath)) {
    try { unlinkSync(plistPath); removed = true; } catch { removed = false; }
  }
  return { launchd_label: label, launchd_plist_path: plistPath, bootout_returncode: bootout.status, bootout_stderr: redact(bootout.stderr || ""), launchd_plist_removed: removed, legacy_admin_bootout_returncode: legacyBootout?.status, legacy_admin_bootout_stderr: legacyBootout ? redact(legacyBootout.stderr || "") : undefined, legacy_admin_plist_removed: legacyRemoved };
}

export function backupPathFor(path: string, action: string): string {
  const stamp = new Date().toISOString().replace(/[:.]/g, "");
  const name = path.split("/").pop() || "workflow.json";
  const dir = `${BACKUP_ROOT}/${stamp}-${safeFileNamePart(action)}`;
  mkdirSync(dir, { recursive: true, mode: 0o700 });
  const dest = `${dir}/${name}.before`;
  copyFileSync(path, dest);
  return dest;
}

export function appendAudit(event: Record<string, unknown>): void {
  mkdirSync(AUDIT_PATH.split("/").slice(0, -1).join("/"), { recursive: true, mode: 0o700 });
  const metadata = redact({
    timestamp_utc: new Date().toISOString(),
    source: "emclaw_m365_connector_workflow_wrapper",
    tokens_logged: false,
    raw_slack_ids_logged: false,
    emails_logged: false,
    graph_called: false,
    slack_send_performed: false,
    m365_content_accessed: false,
    ...event,
  });
  appendFileSync(AUDIT_PATH, `${JSON.stringify(metadata)}\n`, { mode: 0o600 });
}

export function readEmailAlertPolicy(): { max_active_email_alerts_per_user: number; source: string } {
  if (!existsSync(EMAIL_ALERT_POLICY_PATH)) {
    return { max_active_email_alerts_per_user: DEFAULT_EMAIL_ALERT_ACTIVE_QUOTA, source: "default" };
  }
  const policy = objectValue(readJsonFile(EMAIL_ALERT_POLICY_PATH));
  const value = Number(policy.max_active_email_alerts_per_user);
  if (!Number.isInteger(value) || value < MIN_EMAIL_ALERT_ACTIVE_QUOTA || value > MAX_EMAIL_ALERT_ACTIVE_QUOTA) {
    return { max_active_email_alerts_per_user: DEFAULT_EMAIL_ALERT_ACTIVE_QUOTA, source: "invalid_policy_defaulted" };
  }
  return { max_active_email_alerts_per_user: value, source: "policy_file" };
}

export function writeEmailAlertPolicy(maxActive: number, actor: Caller): void {
  mkdirSync(EMAIL_ALERT_POLICY_PATH.split("/").slice(0, -1).join("/"), { recursive: true, mode: 0o700 });
  writePrivateJsonAtomic(EMAIL_ALERT_POLICY_PATH, {
    schema_version: "emclaw.email-alert-policy.v1",
    max_active_email_alerts_per_user: maxActive,
    min_allowed: MIN_EMAIL_ALERT_ACTIVE_QUOTA,
    max_allowed: MAX_EMAIL_ALERT_ACTIVE_QUOTA,
    cadence_floor_seconds: EMAIL_ALERT_MIN_CADENCE_SECONDS,
    cadence_floor_admin_adjustable: false,
    updated_at_utc: isoNow(),
    updated_by_user_ref: actor.user_ref,
    raw_slack_ids_logged: false,
    tokens_logged: false,
    email_content_logged: false,
  });
}

export function rosterDisplayLabel(userRef: string): string {
  if (userRef === "admin") return "Riley/admin";
  try {
    const overlay = objectValue(readJsonFile(PRIVATE_USER_DIRECTORY_PATH));
    const users = Array.isArray(overlay.users) ? overlay.users : [];
    const match = users.find((entry) => objectValue(entry).user_ref === userRef);
    const display = stringValue(objectValue(match).display_name);
    if (display) return display;
  } catch { /* best effort only */ }
  const roster = loadRoster().find((entry) => entry.user_ref === userRef);
  if (roster?.slack_user_hash_short) return `member-${roster.slack_user_hash_short}`;
  return userRef === "admin" ? "Riley/admin" : "member";
}

export function isEmailAlertWorkflowSpec(spec: Record<string, unknown>): boolean {
  const service = canonicalWorkflowService(spec.service);
  const operation = canonicalWorkflowOperation(spec.operation);
  return (service === "mail" || service === "m365.mail") && alertMailOperations.has(operation);
}

export function isActiveWorkflowSpec(spec: Record<string, unknown>): boolean {
  const status = stringValue(spec.activation_status).toLowerCase();
  return spec.enabled === true && status !== "paused" && status !== "terminated" && status !== "quarantined";
}

export function activeEmailAlertWorkflowFilesForOwner(ownerUserRef: string): string[] {
  return listWorkflowFilesForOwner(ownerUserRef).filter((path) => {
    try {
      const spec = readJsonFile(path) as Record<string, unknown>;
      return isEmailAlertWorkflowSpec(spec) && isActiveWorkflowSpec(spec);
    } catch { return false; }
  });
}

export function emailAlertDedupKey(spec: Record<string, unknown>): string {
  const operation = canonicalWorkflowOperation(spec.operation);
  const criteria = objectValue(spec.criteria);
  const filters = objectValue(spec.filters);
  const subject = normalizeSearchText(criteria.subjectContains ?? criteria.subject_contains ?? filters.subjectContains ?? filters.subject_contains ?? spec.subjectContains ?? (spec as Record<string, unknown>).subject_contains);
  const sender = normalizeSearchText(criteria.senderOrDomain ?? criteria.sender_or_domain ?? filters.senderOrDomain ?? filters.sender_or_domain ?? spec.senderOrDomain ?? (spec as Record<string, unknown>).sender_or_domain);
  const keyword = normalizeSearchText(criteria.keyword ?? filters.keyword ?? (spec as Record<string, unknown>).keyword);
  const folderRaw = stringValue(criteria.folderScope ?? criteria.folder_scope ?? filters.folderScope ?? filters.folder_scope ?? (spec as Record<string, unknown>).folderScope);
  const folder = folderRaw === "mailbox_basic" ? "mailbox_basic" : "inbox";
  return [operation, "subj:" + subject, "sender:" + sender, "kw:" + keyword, "folder:" + folder].join("|");
}

export function findActiveEmailAlertDuplicateForOwner(ownerUserRef: string, spec: Record<string, unknown>): { workflow_id: string; workflow_path: string } | null {
  const key = emailAlertDedupKey(spec);
  for (const path of activeEmailAlertWorkflowFilesForOwner(ownerUserRef)) {
    try {
      const existing = readJsonFile(path) as Record<string, unknown>;
      if (emailAlertDedupKey(existing) === key) {
        const wid = stringValue(existing.workflow_id) || (path.split("/").pop() || "").replace(/\.json$/, "");
        return { workflow_id: wid, workflow_path: path };
      }
    } catch { /* ignore unreadable spec */ }
  }
  return null;
}

export function activeEmailAlertCountsByOwner(): Record<string, number> {
  const counts: Record<string, number> = {};
  if (!existsSync(WORKFLOWS_ROOT)) return counts;
  for (const owner of readdirSync(WORKFLOWS_ROOT)) {
    if (owner.startsWith("_")) continue;
    const dir = `${WORKFLOWS_ROOT}/${owner}`;
    try {
      if (!statSync(dir).isDirectory()) continue;
    } catch { continue; }
    const count = activeEmailAlertWorkflowFilesForOwner(owner).length;
    if (count > 0) counts[owner] = count;
  }
  return counts;
}

export function emailAlertQuotaMessage(ownerUserRef: string, countAfterCreate: number, limit: number, clampedText: string, adminTone: boolean): string {
  const remaining = Math.max(0, limit - countAfterCreate);
  const base = adminTone
    ? `Created email alert for ${rosterDisplayLabel(ownerUserRef)}. Active email alerts: ${countAfterCreate} of ${limit}.`
    : `Done — your email alert is set up. You can have up to ${limit} active email alerts at a time. You now have ${countAfterCreate} of ${limit} active.`;
  const tail = remaining > 0
    ? ` You have ${remaining} remaining.`
    : " That's your maximum. Pause or terminate one if you want to add another.";
  return [base, tail, clampedText].filter(Boolean).join(" ");
}

export function emailAlertQuotaRefusal(ownerUserRef: string, count: number, limit: number, adminTone: boolean): string {
  return adminTone
    ? `${rosterDisplayLabel(ownerUserRef)} already has ${count} of ${limit} active email alerts. Pause or terminate one before creating another.`
    : `You already have ${count} of ${limit} active email alerts, which is the maximum. Pause or terminate one of your existing alerts first, then I can set this up.`;
}

export function requestedEmailAlertCadenceSeconds(args: Record<string, unknown>, spec: WorkflowSpec, schedule: Record<string, unknown>): number {
  const trigger = objectValue(args.trigger || spec.trigger);
  const candidates = [
    trigger.seconds,
    args.cadence_seconds,
    spec.cadence_seconds,
    schedule.cadence_seconds,
    args.interval_seconds,
    spec.interval_seconds,
    schedule.interval_seconds,
  ];
  for (const candidate of candidates) {
    const value = Number(candidate);
    if (Number.isFinite(value) && value > 0) return Math.round(value);
  }
  const minuteCandidates = [
    args.cadence_minutes,
    spec.cadence_minutes,
    schedule.cadence_minutes,
    args.interval_minutes,
    spec.interval_minutes,
    schedule.interval_minutes,
    spec.default_alert_cadence_minutes,
  ];
  for (const candidate of minuteCandidates) {
    const value = Number(candidate);
    if (Number.isFinite(value) && value > 0) return Math.round(value * 60);
  }
  if (requestedOneMinuteCadence(args, spec)) return 60;
  return EMAIL_ALERT_MIN_CADENCE_SECONDS;
}

export function applyEmailAlertCadenceFloor(spec: WorkflowSpec, args: Record<string, unknown>, schedule: Record<string, unknown>): { requested_seconds: number; effective_seconds: number; clamped: boolean; message: string } {
  const requested = requestedEmailAlertCadenceSeconds(args, spec, schedule);
  const effective = Math.max(EMAIL_ALERT_MIN_CADENCE_SECONDS, requested);
  const trigger = objectValue(spec.trigger);
  spec.trigger = { ...trigger, kind: "interval", seconds: effective };
  spec.default_alert_cadence_minutes = Math.round(effective / 60);
  spec.cadence_policy = {
    requested_seconds: requested,
    effective_seconds: effective,
    floor_seconds: EMAIL_ALERT_MIN_CADENCE_SECONDS,
    clamped: requested < EMAIL_ALERT_MIN_CADENCE_SECONDS,
    reason: requested < EMAIL_ALERT_MIN_CADENCE_SECONDS ? "hardware_protection_floor" : "requested_cadence_accepted",
  };
  const clamped = requested < EMAIL_ALERT_MIN_CADENCE_SECONDS;
  return {
    requested_seconds: requested,
    effective_seconds: effective,
    clamped,
    message: clamped ? "I set this to check every 3 minutes — that's the fastest cadence allowed to keep the system responsive for everyone." : "",
  };
}

export function parseSpecJson(specJson: unknown): WorkflowSpec {
  if (typeof specJson !== "string" || specJson.trim() === "") throw new Error("spec_json_required");
  const parsed = JSON.parse(specJson) as unknown;
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) throw new Error("spec_json_object_required");
  return parsed as WorkflowSpec;
}

export function stringValue(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}

export function canonicalWorkflowService(value: unknown): string {
  const service = stringValue(value).toLowerCase();
  if (service === "email" || service === "m365.email") return "mail";
  return service;
}

export function canonicalWorkflowOperation(value: unknown): string {
  const operation = stringValue(value).toLowerCase();
  const canonical = operation.startsWith("email.") ? `mail.${operation.slice("email.".length)}` : operation;
  if (/^mail\.(new_(email|message|mail)_alert|email_subject_alert|subject(_contains)?_alert|inbox_alert|message_alert|mail_alert)$/.test(canonical)) {
    return "mail.alert_metadata_delta";
  }
  return canonical;
}

export function objectValue(value: unknown): Record<string, unknown> {
  return value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : {};
}

export function looksLikeSafeLegacyCalendarStatusSpec(args: Record<string, unknown>, spec: WorkflowSpec): boolean {
  const text = JSON.stringify({ args, spec }).toLowerCase();
  const mentionsCalendar = text.includes("calendar");
  const statusOnly = text.includes("status") || text.includes("metadata") || text.includes("m365_status");
  const unsafe = text.includes("mail_send") || text.includes("send_mail") || text.includes("calendar_write") || text.includes("create_event") || text.includes("update_event") || text.includes("delete_event") || text.includes("slack_send\":true") || text.includes("event_bodies\":true") || text.includes("mail_bodies\":true") || text.includes("files\":true");
  return mentionsCalendar && statusOnly && !unsafe;
}

export function calendarWorkflowDefaults(operation: string): { workflowId: string; name: string; schedule: Record<string, unknown>; criteria: Record<string, unknown>; confirmation: string } {
  if (alertCalendarOperations.has(operation) || alertMailOperations.has(operation)) {
    return {
      workflowId: operation === "calendar.important_meeting_alert_metadata" ? "calendar-important-meeting-alert" : "calendar-metadata-alert",
      name: operation === "calendar.important_meeting_alert_metadata" ? "Important meeting metadata alert" : "Calendar metadata alert",
      schedule: { type: "interval", cadence_minutes: 15, timezone: "America/New_York" },
      criteria: {
        intent: "likely_important_meeting",
        subject_keywords: ["client", "investor", "board", "contract", "approval", "deadline", "closing"],
        use_external_attendee_signal: true,
        use_attendee_count_signal: true,
        use_busy_status_signal: true,
        use_time_proximity_signal: true,
      },
      confirmation: "I'll check every 15 minutes for new or changed events matching your alert criteria. I'll use calendar metadata only.",
    };
  }
  if (operation === "calendar.digest_metadata" || operation === "calendar.digest_metadata_delta" || operation === "calendar.daily_brief_metadata") {
    return {
      workflowId: operation === "calendar.daily_brief_metadata" ? "calendar-daily-brief" : "calendar-metadata-digest",
      name: operation === "calendar.daily_brief_metadata" ? "Daily calendar metadata brief" : "Calendar metadata digest",
      schedule: { type: "daily", time: "08:00", timezone: "America/New_York", weekdays_only: false },
      criteria: { intent: "calendar_digest", window: "today", highlight_likely_important: true },
      confirmation: "Your morning calendar brief will look for today's meetings and highlight likely important ones using subject, organizer, time, attendee count, busy status, and external-attendee signals.",
    };
  }
  return {
    workflowId: "daily-calendar-status-summary",
    name: "Daily calendar status summary",
    schedule: { type: "daily", time: "08:00", timezone: "America/New_York" },
    criteria: { intent: "calendar_status_check", window: "today", highlight_likely_important: true },
    confirmation: "I'll keep this as a metadata-only EMClaw workflow status/check with no Slack delivery.",
  };
}


export function mailWorkflowDefaults(operation: string): { workflowId: string; name: string; schedule: Record<string, unknown>; criteria: Record<string, unknown>; confirmation: string } {
  if (operation === "mail.digest_metadata_delta") {
    return {
      workflowId: "email-basic-details-digest",
      name: "Email basic details digest",
      schedule: { type: "daily", time: "08:00", timezone: "America/New_York", weekdays_only: false },
      criteria: { intent: "email_digest", folder: "inbox", window: "today", highlight_likely_important: true },
      confirmation: "Your email digest will check basic inbox details like sender, subject, time, unread status, importance, and attachment flag. It will not open or read emails, previews, attachments, images, or links.",
    };
  }
  return {
    workflowId: operation === "mail.important_email_alert_metadata" ? "important-email-alert" : "email-attention-alert",
    name: operation === "mail.important_email_alert_metadata" ? "Important email alert" : "Email attention alert",
    schedule: { type: "interval", cadence_minutes: 3, timezone: "America/New_York" },
    criteria: {
      intent: "likely_important_email",
      folder: "inbox",
      subject_keywords: ["urgent", "important", "action", "approve", "approval", "contract", "deadline", "client", "customer", "investor", "board", "invoice", "closing", "signature", "review"],
      use_sender_display_signal: true,
      use_importance_flag: true,
      use_unread_status: true,
      use_recency_signal: true,
      use_attachment_flag: true,
      use_safe_preference_rules: true,
    },
    confirmation: "I'll check for new inbox items that likely need attention. I will only check basic email details and will not open or read emails, previews, attachments, images, or links.",
  };
}


export function rawTextIncludes(value: unknown, patterns: RegExp[]): boolean {
  const text = typeof value === "string" ? value : JSON.stringify(value ?? "");
  return patterns.some((pattern) => pattern.test(text));
}

export function requestedExplicitNoSend(args: Record<string, unknown>, spec: Record<string, unknown>): boolean {
  return rawTextIncludes({ delivery: args.delivery ?? spec.delivery, live_effects: args.live_effects ?? spec.live_effects, mode: args.mode ?? spec.mode, prompt: args.prompt ?? spec.prompt, text: args.text ?? spec.text }, [
    /no[-_ ]?send/i,
    /status[-_ ]?check/i,
    /without\s+sending/i,
    /do\s+not\s+send/i,
    /slack_send[\"']?\s*[:=]\s*false/i,
    /delivery[\"']?\s*[:=]\s*[\"']?(none|status_check_no_send|no_send)/i,
  ]);
}

export function requestedActiveSlackMailAlert(args: Record<string, unknown>, spec: Record<string, unknown>, delivery: Record<string, unknown>, isMailWorkflow: boolean, operation: string): boolean {
  if (!isMailWorkflow || !alertMailOperations.has(operation)) return false;
  if (requestedExplicitNoSend(args, spec)) return false;
  const deliveryRaw = args.delivery ?? spec.delivery;
  const scheduleRaw = args.schedule ?? spec.schedule ?? spec.trigger;
  const text = JSON.stringify({ args, spec, deliveryRaw, scheduleRaw });
  return delivery.slack_send === true
    || args.live_effects === true
    || spec.live_effects === true
    || /slack[_ -]?(send|alert|message)|alert\s+me\s+in\s+slack|delivery[\"']?\s*[:=]\s*[\"']?slack/i.test(String(deliveryRaw ?? ""))
    || (/active\s+slack\s+alert|alert\s+me\s+in\s+slack|notify\s+me\s+in\s+slack|notify[^{}]{0,80}slack|slack\s+alert/i.test(text) && /subjectContains|subject[_ -]?containing|senderOrDomain|sender[_ -]?or[_ -]?domain|from|new\s+email|inbox|emails?/i.test(text));
}

export function requestedOneMinuteCadence(args: Record<string, unknown>, spec: Record<string, unknown>): boolean {
  const text = JSON.stringify({ schedule: args.schedule ?? spec.schedule, trigger: args.trigger ?? spec.trigger, cadence: args.cadence ?? spec.cadence, default_alert_cadence_minutes: args.default_alert_cadence_minutes ?? spec.default_alert_cadence_minutes });
  return /\*\/1|every\s+1\s+minute|1\s*min|60\s*sec|cadence_minutes[\"']?\s*[:=]\s*1|seconds[\"']?\s*[:=]\s*60/i.test(text);
}

export function workflowIdForMailAlert(baseId: string, args: Record<string, unknown>, spec: Record<string, unknown>, live: boolean): string {
  if (!live) return baseId;
  const explicit = stringValue(args.workflow_id || args.workflowId || spec.workflow_id || spec.workflowId);
  if (explicit && explicit !== "email-attention-alert" && explicit !== "important-email-alert") return assertWorkflowId(explicit);
  const subject = stringValue(args.subjectContains || args.subject_contains || spec.subjectContains || spec.subject_contains || objectValue(args.criteria).subjectContains || objectValue(spec.criteria).subjectContains);
  const sender = stringValue(args.senderOrDomain || args.sender_or_domain || spec.senderOrDomain || spec.sender_or_domain || objectValue(args.criteria).senderOrDomain || objectValue(args.criteria).sender_or_domain || objectValue(spec.criteria).senderOrDomain || objectValue(spec.criteria).sender_or_domain);
  const basis = subject || sender;
  if (basis) return assertWorkflowId(`${baseId}-${createHash("sha256").update(basis).digest("hex").slice(0, 8)}`);
  return assertWorkflowId(`${baseId}-${new Date().toISOString().replace(/[^0-9T]/g, "").slice(0, 15).toLowerCase()}`);
}

export function mailWorkflowStatePath(ownerUserRef: string, workflowId: string): string {
  return MAIL_WORKFLOW_STATE_ROOT + "/" + safeFileNamePart(ownerUserRef) + "/" + assertWorkflowId(workflowId) + ".json";
}

export function mailWorkflowInitialState(ownerUserRef: string, workflowId: string, operation: string, slackSend = false): Record<string, unknown> {
  const now = new Date().toISOString();
  return {
    schema_version: "m8d.mail-delta-state.v1",
    owner_user_ref: ownerUserRef,
    workflow_id: workflowId,
    service: "m365.mail",
    operation,
    folder: "inbox",
    status: "reserved",
    initial_watermark_policy: "watermark_now_no_backfill_unless_explicitly_requested",
    watermark_utc: now,
    delta_token: null,
    dedupe_strategy: "stable_message_id_and_conversation_id_not_subject",
    processed_message_id_hashes: [],
    processed_conversation_id_hashes: [],
    failed_run_policy: "do_not_advance_watermark_beyond_unprocessed_items",
    restart_durable: true,
    admin_visibility: "metadata_only",
    content_access: "basic_email_details_only",
    body_read: false,
    bodyPreview_read: false,
    uniqueBody_read: false,
    mime_read: false,
    headers_read: false,
    attachments_read: false,
    raw_email_addresses_logged: false,
    mailbox_write: false,
    slack_send: slackSend,
    live_effect_delivery_ready: slackSend,
    created_at_utc: now,
    updated_at_utc: now,
  };
}

export function reserveMailWorkflowState(spec: WorkflowSpec, caller: Caller): string | null {
  const operation = stringValue(spec.operation);
  if (!deltaMailOperations.has(operation)) return null;
  const workflowId = assertWorkflowId(spec.workflow_id);
  const statePath = mailWorkflowStatePath(caller.user_ref, workflowId);
  mkdirSync(statePath.split("/").slice(0, -1).join("/"), { recursive: true, mode: 0o700 });
  if (!existsSync(statePath) || objectValue(spec.delivery).slack_send === true) writePrivateJsonAtomic(statePath, mailWorkflowInitialState(caller.user_ref, workflowId, operation, objectValue(spec.delivery).slack_send === true));
  return statePath;
}

export function calendarWorkflowStatePath(ownerUserRef: string, workflowId: string): string {
  return CALENDAR_WORKFLOW_STATE_ROOT + "/" + safeFileNamePart(ownerUserRef) + "/" + assertWorkflowId(workflowId) + ".json";
}

export function calendarWorkflowInitialState(ownerUserRef: string, workflowId: string, operation: string): Record<string, unknown> {
  const now = new Date().toISOString();
  return {
    schema_version: "m8b.calendar-delta-state.v1",
    owner_user_ref: ownerUserRef,
    workflow_id: workflowId,
    service: "m365.calendar",
    operation,
    status: "reserved",
    initial_watermark_policy: "watermark_now_no_backfill_unless_explicitly_requested",
    watermark_utc: now,
    delta_token: null,
    dedupe_strategy: "stable_event_id_not_subject",
    processed_event_ids_hashes: [],
    failed_run_policy: "do_not_advance_watermark_beyond_unprocessed_items",
    restart_durable: true,
    admin_visibility: "metadata_only",
    content_access: "metadata_only",
    body_read: false,
    attachments_read: false,
    raw_attendee_emails_logged: false,
    meeting_links_logged: false,
    calendar_write: false,
    slack_send: false,
    created_at_utc: now,
    updated_at_utc: now,
  };
}

export function reserveCalendarWorkflowState(spec: WorkflowSpec, caller: Caller): string | null {
  const operation = stringValue(spec.operation);
  if (!deltaCalendarOperations.has(operation)) return null;
  const workflowId = assertWorkflowId(spec.workflow_id);
  const statePath = calendarWorkflowStatePath(caller.user_ref, workflowId);
  mkdirSync(statePath.split("/").slice(0, -1).join("/"), { recursive: true, mode: 0o700 });
  if (!existsSync(statePath)) writePrivateJsonAtomic(statePath, calendarWorkflowInitialState(caller.user_ref, workflowId, operation));
  return statePath;
}

export function reserveMemberWorkflowState(spec: WorkflowSpec, caller: Caller): string | null {
  return reserveCalendarWorkflowState(spec, caller) || reserveMailWorkflowState(spec, caller);
}

export function normalizeWorkflowCreateInput(input: unknown, caller: Caller, adminMode: boolean): WorkflowSpec {
  const args = objectValue(input);
  const spec = parseSpecJson(Object.prototype.hasOwnProperty.call(args, "spec_json") ? args.spec_json : input);
  let appId = stringValue(args.app_id || spec.app_id || spec.appId);
  let service = canonicalWorkflowService(args.service || spec.service);
  let operation = canonicalWorkflowOperation(args.operation || args.operation_template || spec.operation || spec.operation_template);

  if (!adminMode && (!appId || !service || !operation) && looksLikeSafeLegacyCalendarStatusSpec(args, spec)) {
    appId = "microsoft365";
    service = "calendar";
    operation = "calendar.metadata_daily_status_check";
  }

  if (!adminMode) {
    const isCalendarWorkflow = service === "calendar" && allowedMemberCalendarOperations.has(operation);
    const isMailWorkflow = (service === "mail" || service === "m365.mail") && allowedMemberMailOperations.has(operation);
    if (appId !== "microsoft365" || (!isCalendarWorkflow && !isMailWorkflow)) {
      throw new Error("normalized_m365_member_workflow_metadata_required");
    }
    if (stringValue(args.scope || spec.scope) && stringValue(args.scope || spec.scope) !== "single_user") {
      throw new Error("member_workflow_scope_must_be_single_user");
    }
    if (tokenFileExistsByStat(caller.user_ref) !== true) {
      throw new Error("microsoft365_connection_not_ready_by_token_stat");
    }
    spec.app_id = "microsoft365";
    spec.service = isMailWorkflow ? "m365.mail" : "calendar";
    spec.operation = operation;
    spec.operation_template = operation;
    spec.scope = "single_user";
    spec.owner_user_ref = caller.user_ref;
    spec.participants = { mode: "self", selection_kind: "self" };
    spec.data_access_level = isMailWorkflow ? "basic_email_details_only" : "metadata_only";
    spec.content_access_level = isMailWorkflow ? "basic_email_details_only" : "metadata_only";
    spec.live_effect_level = "none";
    spec.calendar_write = false;
    spec.mailbox_write = false;
    spec.body_read = false;
    spec.bodyPreview_read = false;
    spec.uniqueBody_read = false;
    spec.mime_read = false;
    spec.headers_read = false;
    spec.attachments_read = false;
    spec.raw_email_addresses_exposed = false;
    spec.raw_attendee_emails_exposed = false;
    spec.meeting_links_exposed = false;
    spec.approval_required_for = Array.isArray(spec.approval_required_for) ? spec.approval_required_for : [];
    spec.max_run_duration_seconds = typeof spec.max_run_duration_seconds === "number" ? spec.max_run_duration_seconds : 300;
    const defaults = isMailWorkflow ? mailWorkflowDefaults(operation) : calendarWorkflowDefaults(operation);
    spec.enabled = typeof spec.enabled === "boolean" ? spec.enabled : false;
    spec.activation_status = stringValue(spec.activation_status) || "check_only_not_scheduled";
    spec.scheduling_active = false;
    spec.schema_version = stringValue(spec.schema_version) || (isMailWorkflow ? "m8d.mail-workflow.v1" : "m8b.calendar-workflow.v1");
    const delivery = objectValue(args.delivery || spec.delivery);
    const liveSlackIntent = requestedActiveSlackMailAlert(args, spec, delivery, isMailWorkflow, operation);
    if (!stringValue(spec.workflow_id)) spec.workflow_id = workflowIdForMailAlert(defaults.workflowId, args, spec, liveSlackIntent);
    if (!stringValue(spec.name)) spec.name = defaults.name;

    const schedule = { ...defaults.schedule, ...objectValue(args.schedule || spec.schedule) };
    const timezone = stringValue(schedule.timezone || args.timezone || spec.timezone) || "America/New_York";
    if (!spec.trigger || typeof spec.trigger === "string") {
      if (alertCalendarOperations.has(operation) || alertMailOperations.has(operation)) {
        const cadenceMinutes = isMailWorkflow ? (Number(schedule.cadence_minutes || spec.default_alert_cadence_minutes || 3) || 3) : (liveSlackIntent || requestedOneMinuteCadence(args, spec) ? 1 : (Number(schedule.cadence_minutes || spec.default_alert_cadence_minutes || 15) || 15));
        spec.trigger = { kind: "interval", seconds: Math.max(60, Math.round(cadenceMinutes * 60)) };
      } else {
        spec.trigger = {
          kind: stringValue(schedule.type) === "weekly" ? "weekly" : "daily",
          time: stringValue(schedule.time) || "08:00",
          timezone,
          weekdays_only: schedule.weekdays_only === true || /weekday/i.test(JSON.stringify({ args, spec })),
        };
      }
    }
    spec.schedule = schedule;
    spec.timezone = timezone;
    spec.criteria = { ...defaults.criteria, ...objectValue(spec.criteria), ...objectValue(args.criteria) };
    const subjectContains = stringValue(args.subjectContains || args.subject_contains || spec.subjectContains || spec.subject_contains);
    if (isMailWorkflow && subjectContains && !stringValue((spec.criteria as Record<string, unknown>).subjectContains)) {
      (spec.criteria as Record<string, unknown>).subjectContains = subjectContains;
    }
    const senderOrDomain = stringValue(args.senderOrDomain || args.sender_or_domain || spec.senderOrDomain || spec.sender_or_domain);
    if (isMailWorkflow && senderOrDomain && !stringValue((spec.criteria as Record<string, unknown>).senderOrDomain)) {
      (spec.criteria as Record<string, unknown>).senderOrDomain = senderOrDomain;
    }
    spec.output_formatter = stringValue(spec.output_formatter) || (isMailWorkflow ? (alertMailOperations.has(operation) ? "secretary_email_alert_basic_details_formatter" : "secretary_email_digest_basic_details_formatter") : (alertCalendarOperations.has(operation) ? "secretary_calendar_alert_metadata_formatter" : "secretary_calendar_digest_metadata_formatter"));
    spec.member_confirmation_text = stringValue(spec.member_confirmation_text) || defaults.confirmation;

    const requestedLiveSlack = liveSlackIntent || delivery.slack_send === true || args.live_effects === true || spec.live_effects === true;
    const selfMailSlackAlertAllowed = requestedLiveSlack && isSelfOwnedMailSlackAlertSpec({ ...spec, created_by_user_ref: caller.user_ref, delivery: { ...delivery, mode: "slack_send", slack_send: true }, live_effects: true }, caller);
    if (requestedLiveSlack && !selfMailSlackAlertAllowed) {
      throw new Error(isMailWorkflow ? "mail_alert_slack_send_not_enabled" : "workflow_live_effects_not_enabled");
    }
    const contentAccess = objectValue(args.content_access || spec.content_access);
    if (contentAccess.event_bodies === true || contentAccess.mail_bodies === true || contentAccess.files === true || contentAccess.body === true || contentAccess.bodyPreview === true || contentAccess.uniqueBody === true || contentAccess.headers === true || contentAccess.mime === true || contentAccess.raw_email_addresses === true) {
      throw new Error(isMailWorkflow ? "member_mail_workflow_content_bodies_not_allowed" : "member_calendar_workflow_content_bodies_not_allowed");
    }
    spec.delivery = selfMailSlackAlertAllowed ? { ...delivery, mode: "slack_send", slack_send: true } : { mode: "status_check_no_send", ...delivery, slack_send: false };
    let cadencePolicy: { requested_seconds: number; effective_seconds: number; clamped: boolean; message: string } | null = null;
    if (isMailWorkflow && alertMailOperations.has(operation)) {
      cadencePolicy = applyEmailAlertCadenceFloor(spec, args, schedule);
      schedule.cadence_minutes = cadencePolicy.effective_seconds / 60;
    }
    if (selfMailSlackAlertAllowed) {
      spec.member_confirmation_text = ["Active Slack email alert is enabled for your own inbox. I will check every 3 minutes for matching new inbox email using basic email details only, and I will not read email bodies, previews, attachments, images, links, or raw email addresses.", cadencePolicy?.message || ""].filter(Boolean).join(" ");
    } else if (cadencePolicy?.message) {
      spec.member_confirmation_text = [stringValue(spec.member_confirmation_text), cadencePolicy.message].filter(Boolean).join(" ");
    }
    spec.content_access = isMailWorkflow
      ? { ...contentAccess, basic_email_details_only: true, metadata_only: true, body: false, bodyPreview: false, uniqueBody: false, mime: false, internet_headers: false, headers: false, attachments: false, attachment_content: false, images: false, OCR: false, raw_email_addresses: false, recipients: false, mailbox_writes: false, calendar: false, files: false }
      : { ...contentAccess, metadata_only: true, event_bodies: false, calendar_body_description: false, attachments: false, raw_attendee_emails: false, meeting_links: false, mail_bodies: false, files: false };
    spec.allowed_fields = Array.isArray(spec.allowed_fields) ? spec.allowed_fields : (isMailWorkflow ? ["id_or_hashed_redacted_message_id", "subject", "sender_display_or_safe_metadata", "receivedDateTime", "importance", "isRead", "hasAttachments", "parentFolderId_or_safe_folder_label", "conversationId_or_hash", "delta_token_state_private_only"] : ["id_or_hashed_redacted_event_id", "subject", "organizer_display_or_safe_metadata", "start", "end", "showAs_or_busy_free", "responseStatus", "attendee_count", "external_attendee_count_or_flag", "safe_location_label_no_join_url", "onlineMeeting_flag_only_no_join_url", "recurrence_summary_if_safe", "delta_token_state_private_only"]);
    spec.excluded_fields = Array.isArray(spec.excluded_fields) ? spec.excluded_fields : (isMailWorkflow ? ["body", "bodyPreview", "uniqueBody", "MIME", "internetMessageHeaders", "attachments", "attachment_content", "images", "OCR", "raw_email_addresses", "recipients", "mailbox_writes", "calendar_or_file_content"] : ["body", "content", "description", "attachments", "raw_attendee_emails_user_facing", "online_meeting_join_url", "calendar_writes", "event_create_update_delete", "email_body_mail_file_content"]);
    if (deltaCalendarOperations.has(operation) || deltaMailOperations.has(operation)) {
      const workflowIdForState = assertWorkflowId(spec.workflow_id);
      spec.default_alert_cadence_minutes = isMailWorkflow ? 3 : (selfMailSlackAlertAllowed ? 1 : 15);
      if (isMailWorkflow && alertMailOperations.has(operation) && objectValue(spec.trigger).seconds && Number(objectValue(spec.trigger).seconds) < EMAIL_ALERT_MIN_CADENCE_SECONDS) {
        cadencePolicy = applyEmailAlertCadenceFloor(spec, args, schedule);
      }
      if (selfMailSlackAlertAllowed) {
        spec.enabled = true;
        spec.activation_status = "active_scheduled";
        spec.scheduling_active = true;
        spec.live_effects = true;
        spec.live_effect_level = "slack_send";
        spec.runs_as_user_ref = caller.user_ref;
        spec.approval_required_for = Array.from(new Set([...(Array.isArray(spec.approval_required_for) ? spec.approval_required_for.map(String) : []), "live_effects", operation]));
        if (cadencePolicy) spec.trigger = { kind: "interval", seconds: cadencePolicy.effective_seconds };
      }
      spec.delta_watermark = {
        supported: true,
        state_path: isMailWorkflow ? mailWorkflowStatePath(caller.user_ref, workflowIdForState) : calendarWorkflowStatePath(caller.user_ref, workflowIdForState),
        initial_watermark: "now",
        backfill_default: false,
        dedupe_by: isMailWorkflow ? "stable_message_id_and_conversation_id" : "stable_event_id",
        failed_run_policy: "do_not_advance_watermark_beyond_unprocessed_items",
        restart_durable: true,
        admin_visibility: "metadata_only",
      };
      spec.state_path = isMailWorkflow ? mailWorkflowStatePath(caller.user_ref, workflowIdForState) : calendarWorkflowStatePath(caller.user_ref, workflowIdForState);
    }
    if (!Array.isArray(spec.steps) || spec.steps.length === 0) {
      spec.steps = [{
        id: operation,
        actor: caller.user_ref,
        actor_user_ref: caller.user_ref,
        app_id: "microsoft365",
        service: isMailWorkflow ? "m365.mail" : "calendar",
        operation,
        kind: isMailWorkflow ? `m365.mail.${operation}` : `m365.calendar.${operation}`,
        data_access_level: isMailWorkflow ? "basic_email_details_only" : "metadata_only",
        live_effect: selfMailSlackAlertAllowed,
        live_effect_level: selfMailSlackAlertAllowed ? "slack_send" : "none",
        calendar_write: false,
        mailbox_write: false,
        body_read: false,
        bodyPreview_read: false,
        uniqueBody_read: false,
        mime_read: false,
        headers_read: false,
        attachments_read: false,
        raw_email_addresses_exposed: false,
        raw_attendee_emails_exposed: false,
        meeting_links_exposed: false,
      }];
    }
  }
  return spec;
}

export function runExecutor(args: string[]): ToolResult {
  const result = spawnSync(PYTHON_BIN, [WORKFLOW_EXECUTOR_PATH, ...args], { encoding: "utf8", maxBuffer: 1024 * 1024 });
  const stdout = typeof result.stdout === "string" ? result.stdout.trim() : "";
  const stderr = typeof result.stderr === "string" ? result.stderr.trim() : "";
  let parsed: unknown = null;
  if (stdout) {
    try { parsed = JSON.parse(stdout); } catch { parsed = { raw_stdout: redact(stdout) }; }
  }
  return {
    executor_returncode: result.status,
    executor_stdout: redact(parsed),
    executor_stderr: redact(stderr),
    executor_ok: result.status === 0 && Boolean(parsed && typeof parsed === "object" && (parsed as { ok?: unknown }).ok === true),
  };
}

export function validateSpecWithExecutor(spec: WorkflowSpec): ToolResult {
  const tmp = `/private/tmp/emclaw-workflow-validate-${Date.now()}-${Math.random().toString(16).slice(2)}.json`;
  try {
    writeFileSync(tmp, `${JSON.stringify(spec, null, 2)}\n`, { mode: 0o600 });
    return runExecutor(["--workflow-file", tmp, "--dry-run"]);
  } finally {
    try { unlinkSync(tmp); } catch { /* best effort cleanup */ }
  }
}

export function workflowSummary(path: string): ToolResult {
  const spec = readJsonFile(path) as Record<string, unknown>;
  const ownerRef = typeof spec.owner_user_ref === "string" ? spec.owner_user_ref : null;
  const createdByRef = typeof spec.created_by_user_ref === "string" ? spec.created_by_user_ref : null;
  return {
    workflow_id: typeof spec.workflow_id === "string" ? spec.workflow_id : path.split("/").pop()?.replace(/\.json$/, ""),
    name: typeof spec.name === "string" ? redact(spec.name) : null,
    scope: typeof spec.scope === "string" ? spec.scope : null,
    owner_user_ref: ownerRef,
    owner_display_label: ownerRef ? rosterDisplayLabel(ownerRef) : null,
    created_by_user_ref: createdByRef,
    created_by_display_label: createdByRef ? rosterDisplayLabel(createdByRef) : null,
    enabled: Boolean(spec.enabled),
    trigger: redact(spec.trigger ?? null),
    service: typeof spec.service === "string" ? spec.service : null,
    operation: typeof spec.operation === "string" ? spec.operation : null,
    runs_as_user_ref: typeof spec.runs_as_user_ref === "string" ? spec.runs_as_user_ref : null,
    created_by_admin_user_ref: typeof spec.created_by_admin_user_ref === "string" ? spec.created_by_admin_user_ref : null,
    admin_delegated: spec.admin_delegated === true,
    activation_status: typeof spec.activation_status === "string" ? spec.activation_status : null,
    connection_state: typeof (spec.connection_used as { state?: unknown } | undefined)?.state === "string" ? (spec.connection_used as { state?: string }).state : null,
    cadence_policy: redact(spec.cadence_policy ?? null),
  };
}

export function listWorkflowFilesForOwner(ownerUserRef: string): string[] {
  const dir = workflowDirForOwner(ownerUserRef);
  if (!existsSync(dir)) return [];
  return readdirSync(dir).filter((name) => name.endsWith(".json")).map((name) => `${dir}/${name}`);
}

export function listAllWorkflowFiles(): string[] {
  if (!existsSync(WORKFLOWS_ROOT)) return [];
  const out: string[] = [];
  for (const owner of readdirSync(WORKFLOWS_ROOT)) {
    if (owner.startsWith("_")) continue;
    const dir = `${WORKFLOWS_ROOT}/${owner}`;
    try {
      if (!statSync(dir).isDirectory()) continue;
      for (const file of readdirSync(dir)) {
        if (file.endsWith(".json")) out.push(`${dir}/${file}`);
      }
    } catch { /* ignore unreadable workflow dirs */ }
  }
  return out;
}

export function findWorkflowPath(workflowId: string): string | null {
  const safeId = assertWorkflowId(workflowId);
  for (const path of listAllWorkflowFiles()) {
    if (path.endsWith(`/${safeId}.json`)) return path;
    try {
      const spec = readJsonFile(path) as Record<string, unknown>;
      if (spec.workflow_id === safeId) return path;
    } catch { /* skip invalid specs */ }
  }
  return null;
}

export function setWorkflowEnabled(path: string, enabled: boolean, caller: Caller, action: string): ToolResult {
  const spec = readJsonFile(path) as Record<string, unknown>;
  const ownerUserRef = stringValue(spec.owner_user_ref) || path.split("/").slice(-2, -1)[0] || caller.user_ref;
  if (enabled === true && spec.enabled !== true && isEmailAlertWorkflowSpec(spec)) {
    const policy = readEmailAlertPolicy();
    const activeEmailAlertCount = activeEmailAlertWorkflowFilesForOwner(ownerUserRef).length;
    if (activeEmailAlertCount >= policy.max_active_email_alerts_per_user) {
      const message = emailAlertQuotaRefusal(ownerUserRef, activeEmailAlertCount, policy.max_active_email_alerts_per_user, caller.role === "admin");
      appendAudit({ event_type: "email_alert_quota_resume_refused", caller_user_ref: caller.user_ref, owner_user_ref: ownerUserRef, workflow_id: spec.workflow_id, active_email_alert_count: activeEmailAlertCount, max_active_email_alerts_per_user: policy.max_active_email_alerts_per_user, workflow_spec_written: false });
      return errorResult("email_alert_quota_limit_reached", message, { workflow_id: spec.workflow_id, owner_user_ref: ownerUserRef, owner_display_label: rosterDisplayLabel(ownerUserRef), active_email_alert_count: activeEmailAlertCount, max_active_email_alerts_per_user: policy.max_active_email_alerts_per_user, workflow_spec_written: false, pause_or_terminate_to_free_slot: true });
    }
  }
  const backup = backupPathFor(path, action);
  spec.enabled = enabled;
  spec.updated_at_utc = new Date().toISOString();
  writeJsonAtomic(path, spec);
  appendAudit({ event_type: action, caller_user_ref: caller.user_ref, workflow_id: spec.workflow_id, enabled, backup_created: backup });
  return okResult({ action, workflow_id: spec.workflow_id, enabled, backup_created: backup, scheduling_active: false, launchd_modified: false });
}

export function terminateWorkflow(path: string, caller: Caller, action: string): ToolResult {
  const spec = readJsonFile(path) as Record<string, unknown>;
  const backup = backupPathFor(path, action);
  const workflowId = assertWorkflowId(spec.workflow_id ?? path.split("/").pop()?.replace(/\.json$/, ""));
  const launchd = (objectValue(spec.delivery).slack_send === true || spec.live_effect_level === "slack_send") ? uninstallMailDeltaLaunchAgent(workflowId, stringValue(spec.owner_user_ref) || caller.user_ref) : null;
  const stamp = new Date().toISOString().replace(/[:.]/g, "");
  const quarantineDir = `${WORKFLOWS_ROOT}/_terminated/${safeFileNamePart(caller.user_ref)}`;
  ensureDir(quarantineDir);
  const quarantinePath = `${quarantineDir}/${stamp}-${workflowId}.json`;
  renameSync(path, quarantinePath);
  appendAudit({ event_type: action, caller_user_ref: caller.user_ref, workflow_id: workflowId, backup_created: backup, quarantine_path: quarantinePath });
  return okResult({ action, workflow_id: workflowId, backup_created: backup, quarantine_path: quarantinePath, hard_deleted: false, launchd });
}



export type DelegatedMemberTarget = { user_ref: string; agent_id: string; display_label: string; connected: boolean; connection_state: string };
export type DelegatedWorkflowResolution = { target: DelegatedMemberTarget; workflowId: string; workflowPath: string; resolved_from: string; receipt_id?: string; ignored_slack_timestamp?: string };


export function delegatedSelectorTokens(user: RosterUser): string[] {
  const agentId = routeAgentForUserRef(user.user_ref) || "";
  const tokens = [user.user_ref, user.slack_user_hash_short || "", agentId, user.user_ref.replace(/^slack-user-ref-/, "")];
  if (user.user_ref === "slack-user-ref-8d4e38d9a1cf7e80") tokens.push("avery");
  if (user.user_ref === "slack-user-ref-db6b1d5cccad3cbf") tokens.push("cameron");
  return tokens.map((token) => normalizeSearch(token)).filter(Boolean);
}

export function resolveDelegatedMemberTarget(params: Record<string, unknown>): DelegatedMemberTarget {
  const rawSelector = [params.target_member, params.target_member_ref, params.target_user_ref, params.target_agent_id, params.member]
    .find((value) => typeof value === "string" && value.trim()) as string | undefined;
  if (!rawSelector) throw new Error("target_member_required");
  const selector = normalizeSearch(rawSelector);
  const members = loadRoster().filter((user) => user.role === "member" && user.status === "active");
  const matches = members.filter((user) => delegatedSelectorTokens(user).includes(selector));
  const uniqueRefs = Array.from(new Set(matches.map((user) => user.user_ref)));
  if (uniqueRefs.length === 0) throw new Error("target_member_not_found");
  if (uniqueRefs.length > 1) throw new Error("target_member_ambiguous");
  const target = members.find((user) => user.user_ref === uniqueRefs[0]);
  if (!target) throw new Error("target_member_not_found");
  if (rawSelector === target.user_ref) {
    const confirmed = typeof params.confirm_target_user_ref === "string" ? params.confirm_target_user_ref.trim() : "";
    if (confirmed !== target.user_ref) throw new Error("target_user_ref_confirmation_required");
  }
  const connected = tokenFileExistsByStat(target.user_ref);
  return {
    user_ref: target.user_ref,
    agent_id: routeAgentForUserRef(target.user_ref) || "",
    display_label: `member-${target.slack_user_hash_short || safeFileNamePart(target.user_ref)}`,
    connected,
    connection_state: connected ? "connected_token_private" : "staged_pending_member_connection",
  };
}


export function delegatedTargetSelectorPresent(params: Record<string, unknown>): boolean {
  return [params.target_member, params.target_member_ref, params.target_user_ref, params.target_agent_id, params.member]
    .some((value) => typeof value === "string" && value.trim().length > 0);
}

export function resolveOptionalDelegatedMemberTarget(params: Record<string, unknown>): DelegatedMemberTarget | null {
  return delegatedTargetSelectorPresent(params) ? resolveDelegatedMemberTarget(params) : null;
}

export function delegatedMemberTargetFromUserRef(userRef: string): DelegatedMemberTarget {
  const user = loadRoster().find((entry) => entry.user_ref === userRef && entry.role === "member" && entry.status === "active");
  if (!user) throw new Error("target_member_not_found");
  const connected = tokenFileExistsByStat(user.user_ref);
  return {
    user_ref: user.user_ref,
    agent_id: routeAgentForUserRef(user.user_ref) || "",
    display_label: "member-" + (user.slack_user_hash_short || safeFileNamePart(user.user_ref)),
    connected,
    connection_state: connected ? "connected_token_private" : "staged_pending_member_connection",
  };
}

export function isSlackTimestampLikeWorkflowId(value: unknown): boolean {
  if (typeof value !== "string") return false;
  const text = value.trim();
  return /^\d{10}\.\d{6}$/.test(text) || /^\d{13,}$/.test(text);
}

export function readWorkflowAuditRecords(): Record<string, unknown>[] {
  if (!existsSync(AUDIT_PATH)) return [];
  return readFileSync(AUDIT_PATH, "utf8").split(/\r?\n/).filter(Boolean).map((line) => {
    try { return JSON.parse(line) as Record<string, unknown>; } catch { return null; }
  }).filter((record): record is Record<string, unknown> => Boolean(record));
}

export function findUniqueMemberWorkflowById(workflowId: string, preferredTarget: DelegatedMemberTarget | null): DelegatedWorkflowResolution | null {
  if (preferredTarget) {
    const workflowPath = workflowPathForOwner(preferredTarget.user_ref, workflowId);
    if (existsSync(workflowPath)) return { target: preferredTarget, workflowId, workflowPath, resolved_from: "explicit_workflow_id_and_target" };
    return null;
  }
  const matches: DelegatedWorkflowResolution[] = [];
  for (const user of loadRoster().filter((entry) => entry.role === "member" && entry.status === "active")) {
    const target = delegatedMemberTargetFromUserRef(user.user_ref);
    const workflowPath = workflowPathForOwner(user.user_ref, workflowId);
    if (existsSync(workflowPath)) matches.push({ target, workflowId, workflowPath, resolved_from: "explicit_workflow_id_unique_member_match" });
  }
  if (matches.length === 1) return matches[0];
  if (matches.length > 1) throw new Error("workflow_id_resolver_error:workflow_id_ambiguous_across_members");
  return null;
}

export function writeDelegatedWorkflowReceiptIndex(record: Record<string, unknown>): void {
  ensureDir(DELEGATED_WORKFLOW_RECEIPT_INDEX_PATH.split("/").slice(0, -1).join("/"));
  const existing = existsSync(DELEGATED_WORKFLOW_RECEIPT_INDEX_PATH) ? objectValue(readJsonFile(DELEGATED_WORKFLOW_RECEIPT_INDEX_PATH)) : {};
  const callerRef = stringValue(record.caller_user_ref) || "admin";
  const targetRef = stringValue(record.target_member_ref);
  const next: Record<string, unknown> = { ...existing, schema_version: "2026-06-01.m10_3_r4", updated_at_utc: new Date().toISOString() };
  const latestByAdmin = objectValue(next.latest_by_admin);
  latestByAdmin[callerRef] = record;
  next.latest_by_admin = latestByAdmin;
  if (targetRef) {
    const latestByAdminTarget = objectValue(next.latest_by_admin_target);
    latestByAdminTarget[callerRef + ":" + targetRef] = record;
    next.latest_by_admin_target = latestByAdminTarget;
  }
  writePrivateJsonAtomic(DELEGATED_WORKFLOW_RECEIPT_INDEX_PATH, next);
}

export function latestDelegatedWorkflowReceipt(caller: Caller, params: Record<string, unknown>, preferredTarget: DelegatedMemberTarget | null): DelegatedWorkflowResolution | null {
  const targetFilter = preferredTarget?.user_ref;
  const candidates: Record<string, unknown>[] = [];
  const addCandidate = (record: Record<string, unknown>) => {
    const callerRef = stringValue(record.caller_user_ref);
    const targetRef = stringValue(record.target_member_ref);
    const workflowId = stringValue(record.workflow_id);
    if (callerRef !== caller.user_ref || !targetRef || !workflowId) return;
    if (targetFilter && targetRef !== targetFilter) return;
    candidates.push(record);
  };
  if (existsSync(DELEGATED_WORKFLOW_RECEIPT_INDEX_PATH)) {
    const index = objectValue(readJsonFile(DELEGATED_WORKFLOW_RECEIPT_INDEX_PATH));
    const byAdminTarget = objectValue(index.latest_by_admin_target);
    const byAdmin = objectValue(index.latest_by_admin);
    if (targetFilter) addCandidate(objectValue(byAdminTarget[caller.user_ref + ":" + targetFilter]));
    addCandidate(objectValue(byAdmin[caller.user_ref]));
  }
  for (const record of readWorkflowAuditRecords().reverse()) {
    if (record.event_type === "admin_delegated_workflow_created") addCandidate(record);
  }
  for (const record of candidates) {
    const targetRef = stringValue(record.target_member_ref);
    const workflowId = stringValue(record.workflow_id);
    if (!targetRef || !workflowId) continue;
    const target = delegatedMemberTargetFromUserRef(targetRef);
    const workflowPath = workflowPathForOwner(target.user_ref, workflowId);
    if (existsSync(workflowPath)) return { target, workflowId, workflowPath, resolved_from: "latest_delegated_workflow_receipt", receipt_id: stringValue(record.receipt_id) || undefined };
  }
  return null;
}

export function delegatedWorkflowResolverError(action: string, err: unknown): ToolResult {
  const message = err instanceof Error ? err.message : String(err);
  if (message.startsWith("workflow_id_resolver_error:")) {
    return errorResult("workflow_id_resolver_error", message.replace(/^workflow_id_resolver_error:/, ""), { action, phantom_success_prevented: true, persisted_workflow_receipt_required: true });
  }
  return errorResult(action + "_failed", message);
}

export function detectSelfWorkflowTargetConflict(specInput: unknown, caller: Caller): string | null {
  const args = objectValue(specInput);
  const spec = parseSpecJson(Object.prototype.hasOwnProperty.call(args, "spec_json") ? args.spec_json : specInput);
  const targetCarrier = {
    target_member: args.target_member ?? spec.target_member,
    target_member_ref: args.target_member_ref ?? spec.target_member_ref,
    target_user_ref: args.target_user_ref ?? spec.target_user_ref,
    target_agent_id: args.target_agent_id ?? spec.target_agent_id,
    member: args.member ?? spec.member,
    owner_user_ref: args.owner_user_ref ?? spec.owner_user_ref,
    runs_as_user_ref: args.runs_as_user_ref ?? spec.runs_as_user_ref,
    runs_as: args.runs_as ?? spec.runs_as,
    delegated_target: spec.delegated_target,
    participants: spec.participants,
  };
  const targetText = normalizeSearch(JSON.stringify(targetCarrier));
  if (!targetText || targetText === "objectobject") return null;
  for (const user of loadRoster().filter((entry) => entry.role === "member" && entry.status === "active")) {
    if (user.user_ref === caller.user_ref) continue;
    for (const token of delegatedSelectorTokens(user)) {
      if (token && targetText.includes(token)) return user.user_ref;
    }
  }
  const explicitOwner = stringValue(args.owner_user_ref ?? spec.owner_user_ref ?? args.runs_as_user_ref ?? spec.runs_as_user_ref ?? args.target_user_ref ?? spec.target_user_ref);
  if (explicitOwner && explicitOwner !== caller.user_ref && explicitOwner !== "self") return explicitOwner;
  return null;
}

export function applyDelegatedWorkflowSafety(spec: WorkflowSpec, admin: Caller, target: DelegatedMemberTarget): WorkflowSpec {
  const operation = canonicalWorkflowOperation(spec.operation);
  const service = canonicalWorkflowService(spec.service);
  const isMailWorkflow = service === "m365.mail" || service === "mail";
  spec.operation = operation;
  spec.operation_template = operation;
  spec.service = isMailWorkflow ? "m365.mail" : service;
  spec.owner_user_ref = target.user_ref;
  spec.scope = "single_user";
  spec.participants = "self";
  spec.delegated_target = { mode: "admin_authored_member_owned", target_member_ref: target.user_ref };
  spec.created_by_user_ref = admin.user_ref;
  spec.created_by_admin_user_ref = admin.user_ref;
  spec.created_by_role = "admin";
  spec.admin_delegated = true;
  spec.runs_as_user_ref = target.user_ref;
  spec.runs_as = target.user_ref;
  spec.connection_used = { source: "target_member_connected_app_account", user_ref: target.user_ref, state: target.connection_state, admin_connection_used: false };
  spec.execution_tool_surface = { source: "target_member_allowed_tools", admin_tools_exposed: false };
  spec.admin_visibility = "config_status_redacted_audit_receipts_only";
  spec.member_visibility = "owned_workflow_and_receipts";
  spec.member_notice = { status: "recorded_not_sent", reason: "manual_slack_send_forbidden_in_ioc_validation", created_by_admin_user_ref: admin.user_ref };
  spec.content_access_level = isMailWorkflow ? "basic_email_details_only" : "metadata_only";
  spec.data_access_level = spec.content_access_level;
  spec.live_effect_level = "none";
  spec.calendar_write = false;
  spec.mailbox_write = false;
  spec.body_read = false;
  spec.bodyPreview_read = false;
  spec.uniqueBody_read = false;
  spec.mime_read = false;
  spec.headers_read = false;
  spec.attachments_read = false;
  spec.raw_email_addresses_exposed = false;
  spec.raw_attendee_emails_exposed = false;
  spec.meeting_links_exposed = false;
  spec.delivery = { ...objectValue(spec.delivery), mode: "status_check_no_send", slack_send: false };
  spec.activation_status = target.connected ? (stringValue(spec.activation_status) || "check_only_not_scheduled") : "staged_pending_member_connection";
  if (!target.connected) spec.enabled = false;
  spec.scheduling_active = false;
  spec.updated_at_utc = new Date().toISOString();
  if (!spec.created_at_utc) spec.created_at_utc = spec.updated_at_utc;
  if (Array.isArray(spec.steps)) {
    spec.steps = spec.steps.map((step) => ({
      ...objectValue(step),
      actor: target.user_ref,
      actor_user_ref: target.user_ref,
      run_as_user_ref: target.user_ref,
      calendar_write: false,
      mailbox_write: false,
      body_read: false,
      bodyPreview_read: false,
      uniqueBody_read: false,
      mime_read: false,
      headers_read: false,
      attachments_read: false,
      raw_email_addresses_exposed: false,
      raw_attendee_emails_exposed: false,
      meeting_links_exposed: false,
    }));
  }
  if ((deltaCalendarOperations.has(operation) || deltaMailOperations.has(operation)) && stringValue(spec.workflow_id)) {
    const workflowIdForState = assertWorkflowId(spec.workflow_id);
    spec.default_alert_cadence_minutes = isMailWorkflow ? 3 : 15;
    if (isMailWorkflow && alertMailOperations.has(operation)) {
      const trigger = objectValue(spec.trigger);
      const requestedSeconds = Number(trigger.seconds || objectValue(spec.schedule).seconds || (Number(objectValue(spec.schedule).cadence_minutes || spec.default_alert_cadence_minutes || 3) * 60));
      const effectiveSeconds = Math.max(EMAIL_ALERT_MIN_CADENCE_SECONDS, Number.isFinite(requestedSeconds) ? Math.round(requestedSeconds) : EMAIL_ALERT_MIN_CADENCE_SECONDS);
      spec.trigger = { ...trigger, kind: "interval", seconds: effectiveSeconds };
      spec.cadence_policy = {
        workflow_type: "email_alert",
        requested_seconds: requestedSeconds,
        effective_seconds: effectiveSeconds,
        floor_seconds: EMAIL_ALERT_MIN_CADENCE_SECONDS,
        clamped: effectiveSeconds !== requestedSeconds,
      };
    }
    const statePath = isMailWorkflow ? mailWorkflowStatePath(target.user_ref, workflowIdForState) : calendarWorkflowStatePath(target.user_ref, workflowIdForState);
    spec.state_path = statePath;
    spec.delta_watermark = {
      ...objectValue(spec.delta_watermark),
      supported: true,
      state_path: statePath,
      initial_watermark: "now",
      backfill_default: false,
      dedupe_by: isMailWorkflow ? "stable_message_id_and_conversation_id" : "stable_event_id",
      failed_run_policy: "do_not_advance_watermark_beyond_unprocessed_items",
      restart_durable: true,
      admin_visibility: "metadata_only",
    };
  }
  return spec;
}

export function delegatedMailCriteriaFromInput(args: Record<string, unknown>, spec: Record<string, unknown>): Record<string, unknown> {
  const criteria = { ...objectValue(spec.criteria), ...objectValue(args.criteria) };
  const subjectContains = stringValue(args.subjectContains || args.subject_contains || spec.subjectContains || spec.subject_contains || criteria.subjectContains || criteria.subject_contains);
  const senderOrDomain = stringValue(args.senderOrDomain || args.sender_or_domain || spec.senderOrDomain || spec.sender_or_domain || criteria.senderOrDomain || criteria.sender_or_domain);
  const keyword = stringValue(args.keyword || spec.keyword || criteria.keyword);
  const folderScope = stringValue(args.folderScope || args.folder_scope || spec.folderScope || spec.folder_scope || criteria.folderScope || criteria.folder_scope) || "inbox";
  if (subjectContains) criteria.subjectContains = subjectContains;
  if (senderOrDomain) criteria.senderOrDomain = senderOrDomain;
  if (keyword) criteria.keyword = keyword;
  criteria.folderScope = folderScope === "mailbox_basic" ? "mailbox_basic" : "inbox";
  return criteria;
}

export function looksLikeSafeDelegatedMailStatusCheck(args: Record<string, unknown>, spec: Record<string, unknown>, service: string, operation: string): boolean {
  const criteria = delegatedMailCriteriaFromInput(args, spec);
  const text = JSON.stringify({ args, spec, criteria });
  const mailish = service === "mail" || service === "m365.mail" || operation.startsWith("mail.") || /inbox|email|mail/i.test(text);
  const safeStatusCheck = /no[-_ ]?send|status[-_ ]?check|watch|alert|notify|subject|sender|from|inbox/i.test(text);
  const hasSafeCriterion = Boolean(stringValue(criteria.subjectContains || criteria.senderOrDomain || criteria.keyword));
  return mailish && safeStatusCheck && (hasSafeCriterion || operation.startsWith("mail."));
}

export function normalizeDelegatedMailStatusCheckFallback(input: unknown, target: DelegatedMemberTarget): WorkflowSpec | null {
  const args = objectValue(input);
  const sourceSpec = parseSpecJson(Object.prototype.hasOwnProperty.call(args, "spec_json") ? args.spec_json : input);
  const rawOperation = canonicalWorkflowOperation(args.operation || args.operation_template || sourceSpec.operation || sourceSpec.operation_template);
  const service = canonicalWorkflowService(args.service || sourceSpec.service);
  if (!looksLikeSafeDelegatedMailStatusCheck(args, sourceSpec, service, rawOperation)) return null;

  const operation = allowedMemberMailOperations.has(rawOperation) ? rawOperation : "mail.alert_metadata_delta";
  const defaults = mailWorkflowDefaults(operation);
  const criteria = delegatedMailCriteriaFromInput(args, sourceSpec);
  const schedule = { ...defaults.schedule, ...objectValue(sourceSpec.schedule), ...objectValue(args.schedule) };
  const triggerInput = objectValue(args.trigger || sourceSpec.trigger);
  const requestedSeconds = Number(triggerInput.seconds || schedule.seconds || (Number(schedule.cadence_minutes || sourceSpec.default_alert_cadence_minutes || 3) * 60));
  const effectiveSeconds = Math.max(EMAIL_ALERT_MIN_CADENCE_SECONDS, Number.isFinite(requestedSeconds) ? Math.round(requestedSeconds) : EMAIL_ALERT_MIN_CADENCE_SECONDS);
  const spec: WorkflowSpec = {
    ...sourceSpec,
    app_id: "microsoft365",
    service: "m365.mail",
    operation,
    operation_template: operation,
    workflow_id: stringValue(args.workflow_id || args.workflowId || sourceSpec.workflow_id || sourceSpec.workflowId) || workflowIdForMailAlert(defaults.workflowId, args, { ...sourceSpec, criteria }, false),
    name: stringValue(args.name || sourceSpec.name) || defaults.name,
    scope: "single_user",
    owner_user_ref: target.user_ref,
    participants: { mode: "self", selection_kind: "self" },
    criteria: { ...defaults.criteria, ...criteria },
    schedule: { ...schedule, type: "interval", cadence_minutes: effectiveSeconds / 60 },
    trigger: { ...triggerInput, kind: "interval", seconds: effectiveSeconds },
    timezone: stringValue(schedule.timezone || args.timezone || sourceSpec.timezone) || "America/New_York",
    enabled: false,
    activation_status: "check_only_not_scheduled",
    scheduling_active: false,
    schema_version: "m8d.mail-workflow.v1",
    approval_required_for: [],
    max_run_duration_seconds: 300,
    delivery: { ...objectValue(sourceSpec.delivery), mode: "status_check_no_send", slack_send: false },
    live_effects: false,
    live_effect_level: "none",
    data_access_level: "basic_email_details_only",
    content_access_level: "basic_email_details_only",
    default_alert_cadence_minutes: effectiveSeconds / 60,
    cadence_policy: {
      workflow_type: "email_alert",
      requested_seconds: Number.isFinite(requestedSeconds) ? Math.round(requestedSeconds) : effectiveSeconds,
      effective_seconds: effectiveSeconds,
      floor_seconds: EMAIL_ALERT_MIN_CADENCE_SECONDS,
      clamped: Number.isFinite(requestedSeconds) ? effectiveSeconds !== Math.round(requestedSeconds) : false,
    },
    content_access: { basic_email_details_only: true, metadata_only: true, body: false, bodyPreview: false, uniqueBody: false, mime: false, internet_headers: false, headers: false, attachments: false, attachment_content: false, images: false, OCR: false, raw_email_addresses: false, recipients: false, mailbox_writes: false, calendar: false, files: false },
    allowed_fields: ["id_or_hashed_redacted_message_id", "subject", "sender_display_or_safe_metadata", "receivedDateTime", "importance", "isRead", "hasAttachments", "parentFolderId_or_safe_folder_label", "conversationId_or_hash", "delta_token_state_private_only"],
    excluded_fields: ["body", "bodyPreview", "uniqueBody", "MIME", "internetMessageHeaders", "attachments", "attachment_content", "images", "OCR", "raw_email_addresses", "recipients", "mailbox_writes", "calendar_or_file_content"],
    output_formatter: "secretary_email_alert_basic_details_formatter",
    member_confirmation_text: "This delegated no-send status-check workflow watches basic inbox email details only. It will not read bodies, previews, attachments, images, links, or raw email addresses.",
    steps: [{
      id: operation,
      actor: target.user_ref,
      actor_user_ref: target.user_ref,
      run_as_user_ref: target.user_ref,
      app_id: "microsoft365",
      service: "m365.mail",
      operation,
      kind: `m365.mail.${operation}`,
      data_access_level: "basic_email_details_only",
      live_effect: false,
      live_effect_level: "none",
      calendar_write: false,
      mailbox_write: false,
      body_read: false,
      bodyPreview_read: false,
      uniqueBody_read: false,
      mime_read: false,
      headers_read: false,
      attachments_read: false,
      raw_email_addresses_exposed: false,
    }],
  };
  return spec;
}

export function normalizeDelegatedWorkflowCreateInput(input: unknown, admin: Caller, target: DelegatedMemberTarget): WorkflowSpec {
  const targetCaller: Caller = { user_ref: target.user_ref, role: "member", status: "active" };
  let spec: WorkflowSpec;
  if (target.connected) {
    try {
      spec = normalizeWorkflowCreateInput(input, targetCaller, false);
    } catch (err) {
      const fallback = normalizeDelegatedMailStatusCheckFallback(input, target);
      if (!fallback) throw err;
      spec = fallback;
    }
  } else {
    const args = objectValue(input);
    spec = parseSpecJson(Object.prototype.hasOwnProperty.call(args, "spec_json") ? args.spec_json : input);
    const operation = canonicalWorkflowOperation(args.operation || args.operation_template || spec.operation || spec.operation_template);
    const serviceArg = canonicalWorkflowService(args.service || spec.service);
    const isCalendarWorkflow = serviceArg === "calendar" && allowedMemberCalendarOperations.has(operation);
    const isMailWorkflow = (serviceArg === "mail" || serviceArg === "m365.mail") && allowedMemberMailOperations.has(operation);
    if (stringValue(args.app_id || spec.app_id || spec.appId) !== "microsoft365" || (!isCalendarWorkflow && !isMailWorkflow)) throw new Error("delegated_workflow_m365_metadata_operation_required");
    spec.app_id = "microsoft365";
    spec.service = isMailWorkflow ? "m365.mail" : "calendar";
    spec.operation = operation;
    spec.operation_template = operation;
    spec.workflow_id = stringValue(spec.workflow_id) || (isMailWorkflow ? mailWorkflowDefaults(operation).workflowId : calendarWorkflowDefaults(operation).workflowId);
    spec.name = stringValue(spec.name) || (isMailWorkflow ? mailWorkflowDefaults(operation).name : calendarWorkflowDefaults(operation).name);
    spec.enabled = false;
    spec.trigger = spec.trigger || { kind: "interval", seconds: 60 };
  }
  return applyDelegatedWorkflowSafety(spec, admin, target);
}

export function adminWorkflowCreateForMember(caller: Caller, params: Record<string, unknown>): ToolResult {
  try {
    requireAdmin(caller);
    const target = resolveDelegatedMemberTarget(params);
    const spec = normalizeDelegatedWorkflowCreateInput(params, caller, target);
    const workflowId = assertWorkflowId(spec.workflow_id);
    const validation = validateSpecWithExecutor(spec);
    if (validation.executor_ok !== true) return errorResult("delegated_workflow_validation_failed", "Executor dry-run validation rejected delegated workflow spec.", { validation });
    const policy = readEmailAlertPolicy();
    const activeEmailAlertCount = activeEmailAlertWorkflowFilesForOwner(target.user_ref).length;
    const createsActiveEmailAlert = isEmailAlertWorkflowSpec(spec) && isActiveWorkflowSpec(spec);
    if (createsActiveEmailAlert) {
      const dup = findActiveEmailAlertDuplicateForOwner(target.user_ref, spec);
      if (dup) {
        appendAudit({ event_type: "admin_delegated_workflow_create_email_alert_idempotent", caller_user_ref: caller.user_ref, target_member_ref: target.user_ref, existing_workflow_id: dup.workflow_id, requested_workflow_id: workflowId, active_email_alert_count: activeEmailAlertCount, workflow_spec_written: false, duplicate_create_prevented: true });
        return okResult({ action: "admin_workflow_create_for_member", workflow_id: dup.workflow_id, workflow_path: dup.workflow_path, idempotent_existing_alert: true, duplicate_create_prevented: true, workflow_spec_written: false, owner_user_ref: target.user_ref, created_by_user_ref: caller.user_ref, runs_as_user_ref: target.user_ref, operation: "mail.alert_metadata_delta", admin_connection_used_for_member_workflow: false, owner_display_label: rosterDisplayLabel(target.user_ref), email_alert_quota: { owner_user_ref: target.user_ref, owner_display_label: rosterDisplayLabel(target.user_ref), active_count: activeEmailAlertCount, max_active_email_alerts_per_user: policy.max_active_email_alerts_per_user, remaining_slots: Math.max(0, policy.max_active_email_alerts_per_user - activeEmailAlertCount) } });
      }
    }
    if (createsActiveEmailAlert && activeEmailAlertCount >= policy.max_active_email_alerts_per_user) {
      const message = emailAlertQuotaRefusal(target.user_ref, activeEmailAlertCount, policy.max_active_email_alerts_per_user, true);
      appendAudit({ event_type: "email_alert_quota_create_refused", caller_user_ref: caller.user_ref, owner_user_ref: target.user_ref, workflow_id: workflowId, active_email_alert_count: activeEmailAlertCount, max_active_email_alerts_per_user: policy.max_active_email_alerts_per_user, workflow_spec_written: false, delegated_admin_create: true });
      return errorResult("email_alert_quota_limit_reached", message, { workflow_id: workflowId, owner_user_ref: target.user_ref, owner_display_label: rosterDisplayLabel(target.user_ref), active_email_alert_count: activeEmailAlertCount, max_active_email_alerts_per_user: policy.max_active_email_alerts_per_user, workflow_spec_written: false, pause_or_terminate_to_free_slot: true });
    }
    const targetCaller: Caller = { user_ref: target.user_ref, role: "member", status: "active" };
    const reservedStatePath = reserveMemberWorkflowState(spec, targetCaller);
    const dir = workflowDirForOwner(target.user_ref);
    ensureDir(dir);
    const workflowPath = workflowPathForOwner(target.user_ref, workflowId);
    if (existsSync(workflowPath)) return errorResult("workflow_already_exists", "Target member workflow spec already exists; refusing overwrite.", { workflow_id: workflowId, owner_user_ref: target.user_ref });
    writeJsonAtomic(workflowPath, spec);
    const receiptId = createHash("sha256").update(["admin_workflow_create_for_member", target.user_ref, workflowId, workflowPath].join(":")).digest("hex").slice(0, 16);
    const workflowStatus = target.connected ? stringValue(spec.activation_status) || "check_only_not_scheduled" : "staged_pending_member_connection";
    const receiptRecord = { event_type: "admin_delegated_workflow_created", caller_user_ref: caller.user_ref, target_member_ref: target.user_ref, workflow_id: workflowId, workflow_path: workflowPath, runs_as_user_ref: target.user_ref, operation: spec.operation, status: workflowStatus, receipt_id: receiptId, connection_state: target.connection_state, admin_content_access: false, delta_watermark_state_reserved: Boolean(reservedStatePath), state_path: reservedStatePath, created_at_utc: new Date().toISOString() };
    appendAudit(receiptRecord);
    writeDelegatedWorkflowReceiptIndex(receiptRecord);
    return okResult({ action: "admin_workflow_create_for_member", workflow_id: workflowId, workflow_path: workflowPath, validation, owner_user_ref: target.user_ref, created_by_user_ref: caller.user_ref, runs_as_user_ref: target.user_ref, operation: spec.operation, status: workflowStatus, receipt_id: receiptId, persisted_workflow_receipt: true, target_member_connection_used: true, admin_connection_used_for_member_workflow: false, connection_state: target.connection_state, staged_only: !target.connected, delta_watermark_state_reserved: Boolean(reservedStatePath), state_path: reservedStatePath, active_or_staged_state: workflowStatus, member_notice_recorded: true, admin_visibility: spec.admin_visibility, member_visibility: spec.member_visibility, email_alert_quota: createsActiveEmailAlert ? { owner_user_ref: target.user_ref, owner_display_label: rosterDisplayLabel(target.user_ref), active_count: activeEmailAlertCount + 1, max_active_email_alerts_per_user: policy.max_active_email_alerts_per_user, remaining_slots: Math.max(0, policy.max_active_email_alerts_per_user - activeEmailAlertCount - 1) } : null, receipt_resolution: { latest_receipt_indexed: true, use_workflow_id_for_followups: workflowId } });
  } catch (err) {
    return errorResult("admin_workflow_create_for_member_failed", err instanceof Error ? err.message : String(err));
  }
}

export const delegatedEditableKeys = new Set(["schedule", "criteria", "enabled", "delivery", "labels", "description", "name", "timezone", "trigger"]);

export function delegatedPatchFromParams(params: Record<string, unknown>): Record<string, unknown> {
  const patch: Record<string, unknown> = {};
  if (typeof params.patch_json === "string" && params.patch_json.trim()) Object.assign(patch, JSON.parse(params.patch_json));
  for (const key of delegatedEditableKeys) if (Object.prototype.hasOwnProperty.call(params, key)) patch[key] = params[key];
  return patch;
}

export function adminWorkflowUpdateForMember(caller: Caller, params: Record<string, unknown>): ToolResult {
  try {
    requireAdmin(caller);
    const target = resolveDelegatedMemberTarget(params);
    const workflowId = assertWorkflowId(params.workflow_id);
    const workflowPath = workflowPathForOwner(target.user_ref, workflowId);
    if (!existsSync(workflowPath)) return errorResult("workflow_not_found", "Target member workflow was not found.", { workflow_id: workflowId, owner_user_ref: target.user_ref });
    const backup = backupPathFor(workflowPath, "admin_workflow_update_for_member");
    const spec = readJsonFile(workflowPath) as WorkflowSpec;
    const before = redact({ schedule: spec.schedule, criteria: spec.criteria, enabled: spec.enabled, delivery: spec.delivery, labels: spec.labels, description: spec.description, name: spec.name, trigger: spec.trigger });
    const patch = delegatedPatchFromParams(params);
    for (const key of Object.keys(patch)) {
      if (!delegatedEditableKeys.has(key)) throw new Error(`delegated_edit_key_not_allowed:${key}`);
      (spec as Record<string, unknown>)[key] = patch[key];
    }
    applyDelegatedWorkflowSafety(spec, caller, target);
    const validation = validateSpecWithExecutor(spec);
    if (validation.executor_ok !== true) return errorResult("delegated_workflow_validation_failed", "Executor dry-run validation rejected updated delegated workflow spec.", { validation, backup_created: backup });
    writeJsonAtomic(workflowPath, spec);
    const after = redact({ schedule: spec.schedule, criteria: spec.criteria, enabled: spec.enabled, delivery: spec.delivery, labels: spec.labels, description: spec.description, name: spec.name, trigger: spec.trigger });
    appendAudit({ event_type: "admin_delegated_workflow_updated", caller_user_ref: caller.user_ref, target_member_ref: target.user_ref, workflow_id: workflowId, backup_created: backup, before, after, admin_content_access: false });
    return okResult({ action: "admin_workflow_update_for_member", workflow_id: workflowId, owner_user_ref: target.user_ref, runs_as_user_ref: target.user_ref, backup_created: backup, validation, before, after, admin_content_access: false });
  } catch (err) {
    return errorResult("admin_workflow_update_for_member_failed", err instanceof Error ? err.message : String(err));
  }
}

export function delegatedWorkflowPathForAction(caller: Caller, params: Record<string, unknown>): DelegatedWorkflowResolution {
  const preferredTarget = resolveOptionalDelegatedMemberTarget(params);
  const rawWorkflowId = stringValue(params.workflow_id);
  if (rawWorkflowId && !isSlackTimestampLikeWorkflowId(rawWorkflowId)) {
    const workflowId = assertWorkflowId(rawWorkflowId);
    const match = findUniqueMemberWorkflowById(workflowId, preferredTarget);
    if (match) return match;
    if (preferredTarget) return { target: preferredTarget, workflowId, workflowPath: workflowPathForOwner(preferredTarget.user_ref, workflowId), resolved_from: "explicit_workflow_id_and_target_missing" };
    throw new Error("workflow_id_resolver_error:workflow_id_not_found_for_any_member");
  }
  const latest = latestDelegatedWorkflowReceipt(caller, params, preferredTarget);
  if (latest) return { ...latest, ignored_slack_timestamp: rawWorkflowId || undefined };
  if (rawWorkflowId && isSlackTimestampLikeWorkflowId(rawWorkflowId)) {
    throw new Error("workflow_id_resolver_error:slack_timestamp_is_not_a_workflow_id_and_no_persisted_delegated_receipt_was_available");
  }
  throw new Error("workflow_id_resolver_error:no_confirmed_delegated_workflow_receipt_available");
}

export function adminWorkflowChangeForMember(caller: Caller, params: Record<string, unknown>, enabled: boolean, action: string): ToolResult {
  try {
    requireAdmin(caller);
    const { target, workflowId, workflowPath } = delegatedWorkflowPathForAction(caller, params);
    if (!existsSync(workflowPath)) return errorResult("workflow_not_found", "Target member workflow was not found.", { workflow_id: workflowId, owner_user_ref: target.user_ref });
    if (enabled && !target.connected) return errorResult("target_member_connection_not_ready", "Disconnected member workflows remain staged until the member connection is ready.", { workflow_id: workflowId, owner_user_ref: target.user_ref, connection_state: target.connection_state });
    const result = setWorkflowEnabled(workflowPath, enabled, caller, action);
    appendAudit({ event_type: `admin_delegated_${action}`, caller_user_ref: caller.user_ref, target_member_ref: target.user_ref, workflow_id: workflowId, runs_as_user_ref: target.user_ref, admin_content_access: false });
    return { ...result, owner_user_ref: target.user_ref, runs_as_user_ref: target.user_ref, admin_content_access: false };
  } catch (err) {
    return delegatedWorkflowResolverError(action, err);
  }
}

export function terminateWorkflowForOwner(path: string, caller: Caller, action: string, quarantineOwnerRef: string): ToolResult {
  const spec = readJsonFile(path) as Record<string, unknown>;
  const backup = backupPathFor(path, action);
  const workflowId = assertWorkflowId(spec.workflow_id ?? path.split("/").pop()?.replace(/\.json$/, ""));
  const stamp = new Date().toISOString().replace(/[:.]/g, "");
  const quarantineDir = `${WORKFLOWS_ROOT}/_terminated/${safeFileNamePart(quarantineOwnerRef)}`;
  ensureDir(quarantineDir);
  const quarantinePath = `${quarantineDir}/${stamp}-${workflowId}.json`;
  renameSync(path, quarantinePath);
  const launchd = objectValue(spec.delivery).slack_send === true ? uninstallMailDeltaLaunchAgent(workflowId, quarantineOwnerRef) : null;
  const receiptId = createHash("sha256").update([action, quarantineOwnerRef, workflowId, quarantinePath].join(":")).digest("hex").slice(0, 16);
  appendAudit({ event_type: action, caller_user_ref: caller.user_ref, target_member_ref: quarantineOwnerRef, workflow_id: workflowId, backup_created: backup, quarantine_path: quarantinePath, status: "quarantined", receipt_id: receiptId, launchd, admin_content_access: false });
  return okResult({ action, workflow_id: workflowId, owner_user_ref: quarantineOwnerRef, runs_as_user_ref: stringValue(spec.runs_as_user_ref) || quarantineOwnerRef, backup_created: backup, quarantine_path: quarantinePath, hard_deleted: false, status: "quarantined", receipt_id: receiptId, persisted_workflow_receipt: true, launchd, admin_content_access: false });
}

export function adminWorkflowTerminateForMember(caller: Caller, params: Record<string, unknown>): ToolResult {
  try {
    requireAdmin(caller);
    const { target, workflowId, workflowPath } = delegatedWorkflowPathForAction(caller, params);
    if (!existsSync(workflowPath)) return errorResult("workflow_not_found", "Target member workflow was not found.", { workflow_id: workflowId, owner_user_ref: target.user_ref });
    return terminateWorkflowForOwner(workflowPath, caller, "admin_workflow_terminate_for_member", target.user_ref);
  } catch (err) {
    return delegatedWorkflowResolverError("admin_workflow_terminate_for_member", err);
  }
}

export function adminWorkflowInspectForMember(caller: Caller, params: Record<string, unknown>): ToolResult {
  try {
    requireAdmin(caller);
    const { target, workflowId, workflowPath } = delegatedWorkflowPathForAction(caller, params);
    if (!existsSync(workflowPath)) return errorResult("workflow_not_found", "Target member workflow was not found.", { workflow_id: workflowId, owner_user_ref: target.user_ref });
    const spec = readJsonFile(workflowPath) as Record<string, unknown>;
    return okResult({ action: "admin_workflow_inspect_for_member", workflow_id: workflowId, owner_user_ref: target.user_ref, runs_as_user_ref: spec.runs_as_user_ref || target.user_ref, status: stringValue(spec.activation_status) || (spec.enabled === true ? "enabled" : "disabled_or_check_only"), receipt_id: createHash("sha256").update(["admin_workflow_inspect_for_member", target.user_ref, workflowId].join(":")).digest("hex").slice(0, 16), workflow: workflowSummary(workflowPath), delegated_config: redact({ service: spec.service, operation: spec.operation, schedule: spec.schedule, criteria: spec.criteria, delivery: spec.delivery, enabled: spec.enabled, activation_status: spec.activation_status, connection_state: (spec.connection_used as { state?: unknown } | undefined)?.state }), admin_content_access: false, private_run_payloads_returned: false, member_content_exposed: false });
  } catch (err) {
    return delegatedWorkflowResolverError("admin_workflow_inspect_for_member", err);
  }
}

export function createWorkflowSpecForCaller(specInput: unknown, caller: Caller, adminMode: boolean): ToolResult {
  try {
    if (adminMode) requireAdmin(caller); else requireSelfCapable(caller);
    if (!adminMode) {
      const targetConflict = detectSelfWorkflowTargetConflict(specInput, caller);
      if (targetConflict) {
        return errorResult("create_my_workflow_self_only_member_target_blocked", "create_my_workflow is self-only and cannot create, edit, or stage workflows for another member. Admin member-targeted workflow requests must use admin_workflow_create_for_member and a persisted delegated receipt.", {
          action: "create_my_workflow",
          caller_user_ref: caller.user_ref,
          target_member_ref: targetConflict,
          required_tool: caller.role === "admin" ? "admin_workflow_create_for_member" : "member_cannot_create_workflow_for_other_member",
          workflow_spec_written: false,
          workflow_specs_created: false,
          persisted_workflow_receipt: false,
          phantom_success_prevented: true,
        });
      }
    }
    const spec = normalizeWorkflowCreateInput(specInput, caller, adminMode);
    const workflowId = assertWorkflowId(spec.workflow_id);
    spec.created_by_user_ref = caller.user_ref;
    spec.created_by_role = adminMode ? "admin" : caller.role;
    if (adminMode) {
      spec.owner_user_ref = typeof spec.owner_user_ref === "string" && spec.owner_user_ref.trim() ? spec.owner_user_ref.trim() : "admin";
    } else {
      spec.owner_user_ref = caller.user_ref;
      spec.scope = "single_user";
      spec.participants = "self";
    }
    const validation = validateSpecWithExecutor(spec);
    if (validation.executor_ok !== true) return errorResult("workflow_validation_failed", "Executor dry-run validation rejected workflow spec.", { validation });
    const reservedStatePath = adminMode ? null : reserveMemberWorkflowState(spec, caller);
    const writeOwner = adminMode ? "admin" : caller.user_ref;
    const dir = workflowDirForOwner(writeOwner);
    ensureDir(dir);
    const target = workflowPathForOwner(writeOwner, workflowId);
    if (existsSync(target)) return errorResult("workflow_already_exists", "Workflow spec already exists; refusing overwrite.", { workflow_id: workflowId });
    const ownerForQuota = stringValue(spec.owner_user_ref) || writeOwner;
    const policy = readEmailAlertPolicy();
    const activeEmailAlertCount = activeEmailAlertWorkflowFilesForOwner(ownerForQuota).length;
    const createsActiveEmailAlert = isEmailAlertWorkflowSpec(spec) && isActiveWorkflowSpec(spec);
    if (createsActiveEmailAlert) {
      const dup = findActiveEmailAlertDuplicateForOwner(ownerForQuota, spec);
      if (dup) {
        appendAudit({ event_type: adminMode ? "admin_workflow_create_email_alert_idempotent" : "my_workflow_create_email_alert_idempotent", caller_user_ref: caller.user_ref, owner_user_ref: ownerForQuota, existing_workflow_id: dup.workflow_id, requested_workflow_id: workflowId, active_email_alert_count: activeEmailAlertCount, workflow_spec_written: false, duplicate_create_prevented: true });
        const idMsg = caller.role === "admin"
          ? `${rosterDisplayLabel(ownerForQuota)} already has an active email alert with these exact criteria (${dup.workflow_id}); no duplicate was created.`
          : "You already have an active email alert for this exact request, so I kept the existing one instead of creating a duplicate. It's still running and will notify you in Slack.";
        return okResult({ action: adminMode ? "create_workflow" : "create_my_workflow", workflow_id: dup.workflow_id, workflow_path: dup.workflow_path, idempotent_existing_alert: true, duplicate_create_prevented: true, workflow_spec_written: false, owner_user_ref: ownerForQuota, owner_display_label: rosterDisplayLabel(ownerForQuota), operation: "mail.alert_metadata_delta", secretary_confirmation_text: idMsg, email_alert_quota: { owner_user_ref: ownerForQuota, owner_display_label: rosterDisplayLabel(ownerForQuota), active_count: activeEmailAlertCount, max_active_email_alerts_per_user: policy.max_active_email_alerts_per_user, remaining_slots: Math.max(0, policy.max_active_email_alerts_per_user - activeEmailAlertCount) } });
      }
    }
    if (createsActiveEmailAlert && activeEmailAlertCount >= policy.max_active_email_alerts_per_user) {
      const message = emailAlertQuotaRefusal(ownerForQuota, activeEmailAlertCount, policy.max_active_email_alerts_per_user, caller.role === "admin");
      appendAudit({ event_type: "email_alert_quota_create_refused", caller_user_ref: caller.user_ref, owner_user_ref: ownerForQuota, workflow_id: workflowId, active_email_alert_count: activeEmailAlertCount, max_active_email_alerts_per_user: policy.max_active_email_alerts_per_user, workflow_spec_written: false });
      return errorResult("email_alert_quota_limit_reached", message, { workflow_id: workflowId, owner_user_ref: ownerForQuota, owner_display_label: rosterDisplayLabel(ownerForQuota), active_email_alert_count: activeEmailAlertCount, max_active_email_alerts_per_user: policy.max_active_email_alerts_per_user, workflow_spec_written: false, pause_or_terminate_to_free_slot: true });
    }
    writeJsonAtomic(target, spec);
    const liveSelfMailAlert = !adminMode && isSelfOwnedMailSlackAlertSpec(spec, caller);
    const initialDelta = liveSelfMailAlert ? runExecutor(["--workflow-file", target, "--execute-mail-delta-alert"]) : null;
    const launchd = liveSelfMailAlert && initialDelta?.executor_ok === true ? installMailDeltaLaunchAgent(target, spec) : null;
    const countAfterCreate = createsActiveEmailAlert ? activeEmailAlertCount + 1 : activeEmailAlertCount;
    const quotaText = createsActiveEmailAlert ? emailAlertQuotaMessage(ownerForQuota, countAfterCreate, policy.max_active_email_alerts_per_user, "", caller.role === "admin") : "";
    const secretaryConfirmation = [stringValue(spec.member_confirmation_text), quotaText].filter(Boolean).join(" ");
    appendAudit({ event_type: adminMode ? "admin_workflow_created" : "my_workflow_created", caller_user_ref: caller.user_ref, workflow_id: workflowId, workflow_path: target, operation: spec.operation, content_access: spec.content_access_level || "metadata_only", delta_watermark_state_reserved: Boolean(reservedStatePath), state_path: reservedStatePath, live_effect_level: spec.live_effect_level, delivery: spec.delivery, cadence_policy: spec.cadence_policy, email_alert_quota: createsActiveEmailAlert ? { owner_user_ref: ownerForQuota, active_count_before: activeEmailAlertCount, active_count_after: countAfterCreate, max_active_email_alerts_per_user: policy.max_active_email_alerts_per_user } : null, initialDelta, launchd });
    return okResult({ action: adminMode ? "create_workflow" : "create_my_workflow", workflow_id: workflowId, workflow_path: target, validation, role_gate: caller.role, workflow_spec_written: true, operation: spec.operation, content_access: spec.content_access_level || "metadata_only", live_effect_level: spec.live_effect_level, delivery: spec.delivery, cadence_policy: spec.cadence_policy, requested_cadence_seconds: objectValue(spec.cadence_policy).requested_seconds, effective_cadence_seconds: objectValue(spec.cadence_policy).effective_seconds, cadence_clamped: objectValue(spec.cadence_policy).clamped === true, email_alert_quota: createsActiveEmailAlert ? { owner_user_ref: ownerForQuota, owner_display_label: rosterDisplayLabel(ownerForQuota), active_count: countAfterCreate, max_active_email_alerts_per_user: policy.max_active_email_alerts_per_user, remaining_slots: Math.max(0, policy.max_active_email_alerts_per_user - countAfterCreate) } : null, delta_watermark_state_reserved: Boolean(reservedStatePath), state_path: reservedStatePath, secretary_confirmation_text: secretaryConfirmation, initial_delta_run: initialDelta, launchd, one_minute_delta_cadence_supported: false, three_minute_delta_cadence_supported: launchd ? true : undefined });
  } catch (err) {
    return errorResult("create_workflow_failed", err instanceof Error ? err.message : String(err));
  }
}

export function createWorkflowSpec(specInput: unknown, callerUserRef: unknown, adminMode: boolean): ToolResult {
  try {
    const caller = resolveCaller(callerUserRef);
    return createWorkflowSpecForCaller(specInput, caller, adminMode);
  } catch (err) {
    return errorResult("create_workflow_failed", err instanceof Error ? err.message : String(err));
  }
}

export function listMyWorkflowsForCaller(caller: Caller): ToolResult {
  requireSelfCapable(caller);
  return okResult({ action: "list_my_workflows", caller_user_ref: caller.user_ref, workflows: listWorkflowFilesForOwner(caller.user_ref).map(workflowSummary) });
}

export function listMyWorkflows(callerUserRef: unknown): ToolResult {
  try {
    return listMyWorkflowsForCaller(resolveCaller(callerUserRef));
  } catch (err) {
    return errorResult("list_my_workflows_failed", err instanceof Error ? err.message : String(err));
  }
}

export function runMyWorkflowForCaller(caller: Caller, workflowId: unknown): ToolResult {
  try {
    requireSelfCapable(caller);
    const id = assertWorkflowId(workflowId);
    const path = workflowPathForOwner(caller.user_ref, id);
    if (!existsSync(path)) return errorResult("workflow_not_found", "Caller-owned workflow was not found.", { workflow_id: id });
    const dryRun = runExecutor(["--workflow-file", path, "--dry-run"]);
    appendAudit({ event_type: "my_workflow_dry_run", caller_user_ref: caller.user_ref, workflow_id: id, executor_returncode: dryRun.executor_returncode });
    return okResult({ action: "run_my_workflow", workflow_id: id, dry_run_only: true, real_execution_implemented: false, executor: dryRun });
  } catch (err) {
    return errorResult("run_my_workflow_failed", err instanceof Error ? err.message : String(err));
  }
}

export function runMyWorkflow(callerUserRef: unknown, workflowId: unknown): ToolResult {
  try {
    return runMyWorkflowForCaller(resolveCaller(callerUserRef), workflowId);
  } catch (err) {
    return errorResult("run_my_workflow_failed", err instanceof Error ? err.message : String(err));
  }
}

export function myWorkflowStatusForCaller(caller: Caller, workflowId: unknown): ToolResult {
  try {
    requireSelfCapable(caller);
    const id = assertWorkflowId(workflowId);
    const path = workflowPathForOwner(caller.user_ref, id);
    if (!existsSync(path)) return errorResult("workflow_not_found", "Caller-owned workflow was not found.", { workflow_id: id });
    const latestRun = latestRunForWorkflow(id);
    return okResult({ action: "my_workflow_status", workflow_id: id, workflow: workflowSummary(path), latest_run: latestRun, recent_runs: latestRun ? [latestRun] : [], workflow_runs_root: WORKFLOW_RUNS_ROOT });
  } catch (err) {
    return errorResult("my_workflow_status_failed", err instanceof Error ? err.message : String(err));
  }
}

export function myWorkflowStatus(callerUserRef: unknown, workflowId: unknown): ToolResult {
  try {
    return myWorkflowStatusForCaller(resolveCaller(callerUserRef), workflowId);
  } catch (err) {
    return errorResult("my_workflow_status_failed", err instanceof Error ? err.message : String(err));
  }
}

export function changeMyWorkflowForCaller(caller: Caller, workflowId: unknown, enabled: boolean, action: string): ToolResult {
  try {
    requireSelfCapable(caller);
    const path = workflowPathForOwner(caller.user_ref, assertWorkflowId(workflowId));
    if (!existsSync(path)) return errorResult("workflow_not_found", "Caller-owned workflow was not found.", { workflow_id: workflowId });
    return setWorkflowEnabled(path, enabled, caller, action);
  } catch (err) {
    return delegatedWorkflowResolverError(action, err);
  }
}

export function changeMyWorkflow(callerUserRef: unknown, workflowId: unknown, enabled: boolean, action: string): ToolResult {
  try {
    return changeMyWorkflowForCaller(resolveCaller(callerUserRef), workflowId, enabled, action);
  } catch (err) {
    return delegatedWorkflowResolverError(action, err);
  }
}

export function terminateMyWorkflowForCaller(caller: Caller, workflowId: unknown): ToolResult {
  try {
    requireSelfCapable(caller);
    const path = workflowPathForOwner(caller.user_ref, assertWorkflowId(workflowId));
    if (!existsSync(path)) return errorResult("workflow_not_found", "Caller-owned workflow was not found.", { workflow_id: workflowId });
    return terminateWorkflow(path, caller, "terminate_my_workflow");
  } catch (err) {
    return errorResult("terminate_my_workflow_failed", err instanceof Error ? err.message : String(err));
  }
}

export function terminateMyWorkflow(callerUserRef: unknown, workflowId: unknown): ToolResult {
  try {
    return terminateMyWorkflowForCaller(resolveCaller(callerUserRef), workflowId);
  } catch (err) {
    return errorResult("terminate_my_workflow_failed", err instanceof Error ? err.message : String(err));
  }
}

export function listAllAutomations(callerUserRef: unknown): ToolResult {
  try {
    const caller = resolveCaller(callerUserRef);
    requireAdmin(caller);
    return okResult({ action: "list_all_automations", caller_user_ref: caller.user_ref, workflows: listAllWorkflowFiles().map(workflowSummary) });
  } catch (err) {
    return errorResult("list_all_automations_failed", err instanceof Error ? err.message : String(err));
  }
}

export function countJsonFilesRecursively(rootDir: string): number {
  if (!existsSync(rootDir)) return 0;
  let count = 0;
  for (const name of readdirSync(rootDir)) {
    const full = `${rootDir}/${name}`;
    try {
      const stat = statSync(full);
      if (stat.isDirectory()) count += countJsonFilesRecursively(full);
      else if (stat.isFile() && name.endsWith(".json")) count += 1;
    } catch { /* ignore unreadable paths */ }
  }
  return count;
}

export type AdminInventoryDetail = "counts" | "active_records" | "terminated_records" | "roster_roles";

export function workflowInventoryCounts(): ToolResult {
  let activeAdmin = 0;
  let activeMember = 0;
  let manifests = 0;
  if (existsSync(WORKFLOWS_ROOT)) {
    for (const name of readdirSync(WORKFLOWS_ROOT)) {
      const full = `${WORKFLOWS_ROOT}/${name}`;
      try {
        const stat = statSync(full);
        if (stat.isFile() && name.endsWith(".json")) {
          manifests += 1;
        } else if (stat.isDirectory() && !name.startsWith("_")) {
          const fileCount = readdirSync(full).filter((file) => file.endsWith(".json")).length;
          if (name === "admin") activeAdmin += fileCount;
          else activeMember += fileCount;
        }
      } catch { /* ignore unreadable paths */ }
    }
  }
  return {
    active_member_workflow_count: activeMember,
    active_admin_workflow_count: activeAdmin,
    system_manifest_count: manifests,
    terminated_or_quarantined_workflow_count: countJsonFilesRecursively(`${WORKFLOWS_ROOT}/_terminated`),
    m365_operation_manifest_counted_as_active_workflow: false,
  };
}

export const terminatedSearchStopWords = new Set([
  "test",
  "case",
  "workflow",
  "workflows",
  "record",
  "records",
  "terminated",
  "quarantined",
  "quarantine",
  "status",
  "check",
]);

export function normalizeSearch(value: unknown): string {
  return typeof value === "string" ? value.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "") : "";
}

export function normalizedSearchVariants(value: unknown): string[] {
  const normalized = normalizeSearch(value);
  if (!normalized) return [];
  const tokens = normalized.split("-").filter(Boolean);
  const stripped = tokens.filter((token) => !terminatedSearchStopWords.has(token)).join("-");
  const compact = normalized.replace(/-/g, "");
  const strippedCompact = stripped.replace(/-/g, "");
  return Array.from(new Set([normalized, stripped, compact, strippedCompact].filter(Boolean)));
}

export function terminatedRecordMatches(filterVariants: string[], candidates: string[]): boolean {
  if (filterVariants.length === 0) return true;
  const candidateVariants = candidates.flatMap((candidate) => normalizedSearchVariants(candidate));
  return filterVariants.some((filter) => candidateVariants.some((candidate) => candidate.includes(filter) || filter.includes(candidate)));
}

export function workflowIdFromTerminatedName(fileName: string): string {
  return fileName.replace(/^\d{4}-\d{2}-\d{2}T\d+Z-/, "").replace(/\.json$/, "");
}

export function timestampFromTerminatedName(fileName: string): string | undefined {
  const match = fileName.match(/^(\d{4}-\d{2}-\d{2}T\d+)Z-/);
  return match ? `${match[1]}Z` : undefined;
}

export function recentWorkflowAuditEvents(): Record<string, unknown>[] {
  if (!existsSync(AUDIT_PATH)) return [];
  return readFileSync(AUDIT_PATH, "utf8").split(/\r?\n/).filter(Boolean).map((line) => {
    try { return JSON.parse(line) as Record<string, unknown>; } catch { return {}; }
  });
}

export function adminTerminatedRecords(params: Record<string, unknown>): ToolResult {
  const rawFilter = params.search ?? params.workflow_id ?? params.name ?? params.query;
  const filter = normalizeSearch(rawFilter);
  const filterVariants = normalizedSearchVariants(rawFilter);
  const audits = recentWorkflowAuditEvents();
  const records: Record<string, unknown>[] = [];
  const root = `${WORKFLOWS_ROOT}/_terminated`;
  if (existsSync(root)) {
    for (const owner of readdirSync(root)) {
      const ownerDir = `${root}/${owner}`;
      try {
        if (!statSync(ownerDir).isDirectory()) continue;
        for (const file of readdirSync(ownerDir).filter((name) => name.endsWith(".json"))) {
          const workflowId = workflowIdFromTerminatedName(file);
          const audit = audits.slice().reverse().find((entry) =>
            entry.event_type === "terminate_my_workflow" &&
            entry.workflow_id === workflowId &&
            typeof entry.quarantine_path === "string" &&
            (entry.quarantine_path as string).endsWith(`/${owner}/${file}`)
          );
          const candidates = [
            workflowId,
            file,
            file.replace(/\.json$/, ""),
            typeof audit?.workflow_id === "string" ? audit.workflow_id : "",
            typeof audit?.quarantine_path === "string" ? audit.quarantine_path.split("/").pop() ?? "" : "",
          ];
          if (!terminatedRecordMatches(filterVariants, candidates)) continue;
          records.push(redact({
            owner_ref: owner,
            workflow_id: workflowId,
            terminated_or_quarantined_timestamp: typeof audit?.timestamp_utc === "string" ? audit.timestamp_utc : timestampFromTerminatedName(file),
            quarantine_present: true,
            backup_present: typeof audit?.backup_created === "string" && (audit.backup_created as string).trim() !== "",
            audit_event_present: Boolean(audit),
            raw_path_exposed: false,
          }) as Record<string, unknown>);
        }
      } catch { /* ignore unreadable owner dirs */ }
    }
  }
  return {
    detail: "terminated_records",
    search: filter || null,
    record_count: records.length,
    records,
    raw_slack_ids_exposed: false,
    emails_exposed: false,
    tokens_exposed: false,
    m365_content_exposed: false,
  };
}

export function latestRunForWorkflow(workflowId: string): Record<string, unknown> | null {
  const roots = existsSync(WORKFLOW_RUNS_ROOT) ? readdirSync(WORKFLOW_RUNS_ROOT).sort().reverse() : [];
  for (const day of roots) {
    const dir = `${WORKFLOW_RUNS_ROOT}/${day}`;
    try {
      if (!statSync(dir).isDirectory()) continue;
      const files = readdirSync(dir)
        .filter((name) => name.startsWith(`${workflowId}-`) && name.endsWith(".json"))
        .sort((a, b) => Number((statSync(`${dir}/${b}`) as { mtimeMs?: number }).mtimeMs || 0) - Number((statSync(`${dir}/${a}`) as { mtimeMs?: number }).mtimeMs || 0));
      for (const file of files) {
        try {
          const path = `${dir}/${file}`;
          const run = readJsonFile(path) as Record<string, unknown>;
          return redact({
            run_id: run.run_id,
            run_record_path: path,
            created_at_utc: run.created_at_utc,
            graph_called: run.graph_called === true,
            final_http_statuses: run.final_http_statuses,
            fetched_count: run.fetched_count,
            matched_count: run.matched_count,
            delta_state_advanced: run.delta_state_advanced === true,
            slack_send_performed: run.slack_send_performed === true,
            body_read: false,
            bodyPreview_read: false,
            attachments_read: false,
            raw_email_addresses_exposed: false,
            mailbox_write: false,
          }) as Record<string, unknown>;
        } catch { /* ignore malformed run */ }
      }
    } catch { /* ignore unreadable run dirs */ }
  }
  return null;
}

export function adminActiveRecords(params: Record<string, unknown>): ToolResult {
  const rawFilter = params.search ?? params.workflow_id ?? params.name ?? params.query;
  const filterVariants = normalizedSearchVariants(rawFilter);
  const records: Record<string, unknown>[] = [];
  for (const path of listAllWorkflowFiles()) {
    try {
      const spec = readJsonFile(path) as Record<string, unknown>;
      const workflowId = stringValue(spec.workflow_id || path.split("/").pop()?.replace(/\.json$/, ""));
      const criteria = objectValue(spec.criteria);
      const subject = stringValue(spec.subjectContains || spec.subject_contains || criteria.subjectContains || criteria.subject_contains);
      const senderOrDomain = stringValue(spec.senderOrDomain || spec.sender_or_domain || criteria.senderOrDomain || criteria.sender_or_domain);
      const candidates = [
        workflowId,
        stringValue(spec.name),
        stringValue(spec.service),
        stringValue(spec.operation),
        subject,
        senderOrDomain,
        ...((Array.isArray(criteria.subject_keywords) ? criteria.subject_keywords : []) as unknown[]).map(stringValue),
      ];
      if (!terminatedRecordMatches(filterVariants, candidates)) continue;
      const summary = workflowSummary(path) as Record<string, unknown>;
      records.push(redact({
        ...summary,
        owner_user_ref: spec.owner_user_ref,
        status: stringValue(spec.activation_status) || (spec.enabled === true ? "enabled" : "disabled_or_check_only"),
        scheduling_active: spec.scheduling_active === true,
        delivery_mode: stringValue(objectValue(spec.delivery).mode),
        slack_send: objectValue(spec.delivery).slack_send === true,
        live_effect_level: stringValue(spec.live_effect_level),
        cadence_seconds: objectValue(spec.trigger).seconds,
        subject_filter_present: Boolean(subject),
        subject_filter_value: subject || null,
        latest_run: latestRunForWorkflow(workflowId),
        raw_paths_exposed: false,
      }) as Record<string, unknown>);
    } catch { /* ignore invalid workflow specs */ }
  }
  return {
    detail: "active_records",
    search: normalizeSearch(rawFilter) || null,
    record_count: records.length,
    records,
    raw_slack_ids_exposed: false,
    emails_exposed: false,
    tokens_exposed: false,
    m365_content_exposed: false,
  };
}

export function readAppConnectionSummary(userRef: string): Record<string, unknown> {
  const path = `${STATE_REDACTED_USERS_ROOT}/${safeFileNamePart(userRef)}/app-connections/app-connections.json`;
  if (!existsSync(path)) return { microsoft365_state: "unknown", slack_state: "unknown" };
  try {
    const data = readJsonFile(path) as Record<string, unknown>;
    const connections = objectValue(data.connections);
    const m365 = objectValue(connections.microsoft365);
    const slack = objectValue(connections.slack);
    return redact({
      microsoft365_state: typeof m365.state === "string" ? m365.state : "unknown",
      microsoft365_evidence_present: Boolean(m365.evidence_present),
      slack_state: typeof slack.state === "string" ? slack.state : "unknown",
      raw_identifiers_stored: data.raw_identifiers_stored === true,
    }) as Record<string, unknown>;
  } catch {
    return { microsoft365_state: "unknown", slack_state: "unknown" };
  }
}

export function routeAgentForUserRef(userRef: string): string | undefined {
  if (userRef === "admin") return "emclaw-admin-v2";
  const path = `${STATE_REDACTED_USERS_ROOT}/${safeFileNamePart(userRef)}/app-connections/app-connections.json`;
  if (existsSync(path)) {
    try {
      const data = readJsonFile(path) as Record<string, unknown>;
      const m365 = objectValue(objectValue(data.connections).microsoft365);
      if (typeof m365.agent_id === "string") return m365.agent_id;
      if (typeof data.agent_id === "string") return data.agent_id;
    } catch { /* ignore malformed app state */ }
  }
  return undefined;
}

export function adminRosterRoles(): ToolResult {
  const users = loadRoster().map((user) => {
    const agentId = routeAgentForUserRef(user.user_ref);
    return redact({
      user_ref: user.user_ref,
      role: user.role,
      status: user.status,
      agent_id: agentId ?? null,
      route_status: agentId ? "mapped" : "unknown",
      connection_state: readAppConnectionSummary(user.user_ref),
      raw_slack_ids_exposed: false,
    }) as Record<string, unknown>;
  });
  return {
    detail: "roster_roles",
    user_count: users.length,
    member_count: users.filter((user) => user.role === "member").length,
    admin_count: users.filter((user) => user.role === "admin").length,
    users,
    raw_slack_ids_exposed: false,
    emails_exposed: false,
    tokens_exposed: false,
    m365_content_exposed: false,
  };
}

export function adminWorkflowInventory(caller: Caller, params: Record<string, unknown> = {}): ToolResult {
  try {
    requireAdmin(caller);
    const detail = typeof params.detail === "string" ? params.detail as AdminInventoryDetail : "counts";
    const detailPayload = detail === "active_records"
      ? adminActiveRecords(params)
      : detail === "terminated_records"
        ? adminTerminatedRecords(params)
        : detail === "roster_roles"
        ? adminRosterRoles()
        : { detail: "counts", ...workflowInventoryCounts() };
    return okResult({
      action: "admin_workflow_inventory",
      caller_user_ref: caller.user_ref,
      ...detailPayload,
      graph_called: false,
      oauth_performed: false,
      token_contents_read: false,
      m365_content_accessed: false,
      workflow_specs_created: false,
      workflow_specs_deleted_or_quarantined: false
    });
  } catch (err) {
    return errorResult("admin_workflow_inventory_failed", err instanceof Error ? err.message : String(err));
  }
}

export function emailAlertQuotaStatusPayload(): Record<string, unknown> {
  const policy = readEmailAlertPolicy();
  const counts = activeEmailAlertCountsByOwner();
  return {
    max_active_email_alerts_per_user: policy.max_active_email_alerts_per_user,
    policy_source: policy.source,
    cadence_floor_seconds: EMAIL_ALERT_MIN_CADENCE_SECONDS,
    cadence_floor_admin_adjustable: false,
    active_counts_by_owner: Object.fromEntries(Object.entries(counts).map(([owner, count]) => [rosterDisplayLabel(owner), count])),
    raw_user_refs_exposed: false,
    token_contents_read: false,
    m365_content_accessed: false,
  };
}

export function adminEmailAlertQuotaGet(caller: Caller): ToolResult {
  try {
    requireAdmin(caller);
    return okResult({
      action: "admin_email_alert_quota_get",
      caller_user_ref: caller.user_ref,
      ...emailAlertQuotaStatusPayload(),
      workflow_specs_created: false,
      workflow_specs_mutated: false,
    });
  } catch (err) {
    return errorResult("admin_email_alert_quota_get_failed", err instanceof Error ? err.message : String(err));
  }
}

export function myEmailAlertQuotaForCaller(caller: Caller): ToolResult {
  try {
    requireSelfCapable(caller);
    const policy = readEmailAlertPolicy();
    const activeCount = activeEmailAlertWorkflowFilesForOwner(caller.user_ref).length;
    const remaining = Math.max(0, policy.max_active_email_alerts_per_user - activeCount);
    return okResult({
      action: "my_email_alert_quota",
      caller_user_ref: caller.user_ref,
      max_active_email_alerts_per_user: policy.max_active_email_alerts_per_user,
      active_email_alert_count: activeCount,
      remaining_slots: remaining,
      at_limit: remaining <= 0,
      policy_source: policy.source,
      cadence_floor_seconds: EMAIL_ALERT_MIN_CADENCE_SECONDS,
      cadence_floor_admin_adjustable: false,
      scope: "caller_own_active_email_alerts_only",
      other_users_counts_exposed: false,
      raw_user_refs_exposed: false,
      graph_calls_performed: false,
      oauth_started: false,
      token_contents_read: false,
      m365_content_accessed: false,
      workflow_specs_created: false,
      workflow_specs_mutated: false,
    });
  } catch (err) {
    return errorResult("my_email_alert_quota_failed", err instanceof Error ? err.message : String(err));
  }
}

export function adminEmailAlertQuotaSet(caller: Caller, params: Record<string, unknown>): ToolResult {
  try {
    requireAdmin(caller);
    const raw = params.max_active_email_alerts_per_user ?? params.max_active ?? params.max ?? params.value;
    const next = Number(raw);
    if (!Number.isInteger(next) || next < MIN_EMAIL_ALERT_ACTIVE_QUOTA || next > MAX_EMAIL_ALERT_ACTIVE_QUOTA) {
      return errorResult("email_alert_quota_out_of_bounds", `Email alert quota must be a whole number from ${MIN_EMAIL_ALERT_ACTIVE_QUOTA} to ${MAX_EMAIL_ALERT_ACTIVE_QUOTA}.`, {
        min_allowed: MIN_EMAIL_ALERT_ACTIVE_QUOTA,
        max_allowed: MAX_EMAIL_ALERT_ACTIVE_QUOTA,
        requested_value: typeof raw === "number" || typeof raw === "string" ? raw : null,
        workflow_specs_mutated: false,
      });
    }
    const oldPolicy = readEmailAlertPolicy();
    const countsBefore = activeEmailAlertCountsByOwner();
    writeEmailAlertPolicy(next, caller);
    const overLimit = Object.entries(countsBefore)
      .filter(([, count]) => count > next)
      .map(([owner, count]) => ({ owner_display_label: rosterDisplayLabel(owner), active_email_alert_count: count }));
    appendAudit({
      event_type: "email_alert_quota_policy_updated",
      caller_user_ref: caller.user_ref,
      old_max_active_email_alerts_per_user: oldPolicy.max_active_email_alerts_per_user,
      new_max_active_email_alerts_per_user: next,
      over_limit_owner_count: overLimit.length,
      existing_workflows_auto_terminated: false,
      cadence_floor_changed: false,
    });
    const note = overLimit.length > 0
      ? `Set max to ${next}. Existing workflows keep running, but ${overLimit.length} owner(s) cannot add new alerts until under the limit.`
      : `Set company-wide max active email alerts per user to ${next}.`;
    return okResult({
      action: "admin_email_alert_quota_set",
      caller_user_ref: caller.user_ref,
      old_max_active_email_alerts_per_user: oldPolicy.max_active_email_alerts_per_user,
      new_max_active_email_alerts_per_user: next,
      message: note,
      over_limit_owners: overLimit,
      existing_workflows_auto_terminated: false,
      cadence_floor_seconds: EMAIL_ALERT_MIN_CADENCE_SECONDS,
      cadence_floor_admin_adjustable: false,
      token_contents_read: false,
      m365_content_accessed: false,
      workflow_specs_auto_terminated: false,
    });
  } catch (err) {
    return errorResult("admin_email_alert_quota_set_failed", err instanceof Error ? err.message : String(err));
  }
}

export function inspectAutomation(callerUserRef: unknown, workflowId: unknown): ToolResult {
  try {
    const caller = resolveCaller(callerUserRef);
    requireAdmin(caller);
    const path = findWorkflowPath(assertWorkflowId(workflowId));
    if (!path) return errorResult("workflow_not_found", "Workflow was not found.", { workflow_id: workflowId });
    const id = assertWorkflowId(workflowId);
    const spec = redact(readJsonFile(path));
    const latestRun = latestRunForWorkflow(id);
    return okResult({ action: "inspect_automation", workflow_id: id, workflow_path: path, spec, latest_run: latestRun, recent_runs: latestRun ? [latestRun] : [] });
  } catch (err) {
    return errorResult("inspect_automation_failed", err instanceof Error ? err.message : String(err));
  }
}

export function changeAutomation(callerUserRef: unknown, workflowId: unknown, enabled: boolean, action: string): ToolResult {
  try {
    const caller = resolveCaller(callerUserRef);
    requireAdmin(caller);
    const path = findWorkflowPath(assertWorkflowId(workflowId));
    if (!path) return errorResult("workflow_not_found", "Workflow was not found.", { workflow_id: workflowId });
    return setWorkflowEnabled(path, enabled, caller, action);
  } catch (err) {
    return delegatedWorkflowResolverError(action, err);
  }
}

export function terminateAutomation(callerUserRef: unknown, workflowId: unknown): ToolResult {
  try {
    const caller = resolveCaller(callerUserRef);
    requireAdmin(caller);
    const path = findWorkflowPath(assertWorkflowId(workflowId));
    if (!path) return errorResult("workflow_not_found", "Workflow was not found.", { workflow_id: workflowId });
    return terminateWorkflow(path, caller, "terminate_automation");
  } catch (err) {
    return errorResult("terminate_automation_failed", err instanceof Error ? err.message : String(err));
  }
}

export function approvePendingAction(callerUserRef: unknown, runId: unknown, approvalToken: unknown, decision: unknown): ToolResult {
  try {
    const caller = resolveCaller(callerUserRef);
    requireAdmin(caller);
    return okResult({
      action: "approve_pending_action",
      caller_user_ref: caller.user_ref,
      run_id: typeof runId === "string" ? redact(runId) : null,
      decision: typeof decision === "string" ? decision : null,
      approval_token_received: typeof approvalToken === "string" && approvalToken.length > 0,
      status: "approval_flow_not_implemented_yet",
      hidden_or_internal_supported: false,
    });
  } catch (err) {
    return errorResult("approve_pending_action_failed", err instanceof Error ? err.message : String(err));
  }
}

export function m365HealthAllConnectedAccounts(callerUserRef: unknown): ToolResult {
  try {
    const caller = resolveCaller(callerUserRef);
    requireAdmin(caller);
    const result = spawnSync(PYTHON_BIN, [M365_HEALTH_HELPER_PATH, "--run"], { encoding: "utf8", maxBuffer: 1024 * 1024 });
    const stdout = typeof result.stdout === "string" ? result.stdout.trim() : "";
    const stderr = typeof result.stderr === "string" ? result.stderr.trim() : "";
    let parsed: unknown = null;
    if (stdout) {
      try { parsed = JSON.parse(stdout); } catch { parsed = { raw_stdout: redact(stdout) }; }
    }
    return okResult({
      action: "m365_health_all_connected_accounts",
      caller_user_ref: caller.user_ref,
      admin_only: true,
      helper_returncode: result.status,
      helper_stdout: redact(parsed),
      helper_stderr: redact(stderr),
      graph_called: true,
      graph_endpoint_called: "/v1.0/me",
      token_contents_read_for_graph_call_only: true,
      no_mail_calendar_drive_calls: true,
      no_m365_content_bodies: true,
      token_values_printed: false,
      token_contents_printed: false,
      ok: result.status === 0 && Boolean(parsed && typeof parsed === "object" && (parsed as { ok?: unknown }).ok === true),
    });
  } catch (err) {
    return errorResult("m365_health_all_connected_accounts_failed", err instanceof Error ? err.message : String(err));
  }
}

