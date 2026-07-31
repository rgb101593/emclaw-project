import { appendFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";

import {
  EMAIL_QUERY_AUDIT_PATH, M365_CLIENT_CONFIG_PATH, PRIVATE_USER_DIRECTORY_PATH,
  ROSTER_PATH, TOKEN_DIR,
} from "./config.js";
import {
  accessTokenFromLoad, activePreferenceValue, addCalendarDays, datedPartsInTimeZone,
  graphMailboxTimeZoneUrl, loadTokenWrapperForGraph, normalizeTimezoneName,
  parseExplicitDateRange, previousMondayParts, refreshAccessToken, safeCalendarText,
  timezoneAwareArgs, userTimeZoneToIana, zonedDayStartInstant,
  type GraphFetchResult,
} from "./calendar.js";
import { loadRoster, requireSelfCapable, sha256Hex } from "./identity.js";
import { ensurePreferenceRecord, preferencePathFor } from "./preferences.js";
import {
  ensureDir, errorResult, isoNow, okResult, readJsonFile, redact,
  tokenFileExistsByStat, writePrivateJsonAtomic,
} from "./util.js";
import type {
  Caller, EmailFolderScope, EmailMetadataMessage, EmailQueryFilters,
  EmailQueryResolvedWindow, EmailQueryWindow, EmailQuestionType, ToolResult,
} from "./types.js";

export const emailQuerySelectFields = ["id", "subject", "from", "receivedDateTime", "importance", "isRead", "hasAttachments", "parentFolderId", "conversationId"];
export const importantEmailKeywords = ["urgent", "important", "action", "approve", "approval", "contract", "deadline", "client", "customer", "investor", "board", "invoice", "closing", "signature", "review"];

export function emailTimezoneFromArgs(args: Record<string, unknown>): string {
  const raw = typeof args.timezone === "string" ? args.timezone.trim() : "";
  return raw || "America/New_York";
}

export function emailWindowFromArgs(args: Record<string, unknown>): EmailQueryResolvedWindow {
  const now = new Date();
  const timezone = emailTimezoneFromArgs(args);
  const dateText = typeof args.dateText === "string" ? args.dateText.trim() : "";
  const question = typeof args.question === "string" ? args.question.trim() : "";
  const rawStart = typeof args.startDateTime === "string" ? args.startDateTime.trim() : "";
  const rawEnd = typeof args.endDateTime === "string" ? args.endDateTime.trim() : "";
  if (rawStart && rawEnd) {
    return { window: "explicit_range", start: new Date(rawStart).toISOString(), end: new Date(rawEnd).toISOString(), date_text: dateText || undefined, timezone, resolution: "explicit_start_end" };
  }
  const parsed = parseExplicitDateRange([dateText, question].filter(Boolean).join(" "), now, timezone);
  if (parsed) return { window: "explicit_range", start: parsed.start.toISOString(), end: parsed.end.toISOString(), date_text: dateText || question || undefined, timezone, resolution: parsed.resolution };
  const wantsLastSenderLookup = /\b(last time|last email|emailed me|from)\b/i.test(question);
  if (!rawStart && !rawEnd && !parsed && wantsLastSenderLookup) {
    const start = new Date(now);
    start.setTime(now.getTime() - 30 * 24 * 60 * 60 * 1000);
    return { window: "explicit_range", start: start.toISOString(), end: now.toISOString(), date_text: dateText || question || undefined, timezone, resolution: "last_30_days_sender_lookup" };
  }
  const raw = typeof args.window === "string" ? args.window.trim().toLowerCase() : "today";
  const window = (["today", "yesterday", "since_yesterday", "last_24_hours", "last_48_hours", "last_7_days", "last_30_days", "this_week", "last_week"].includes(raw) ? raw : "today") as EmailQueryWindow;
  // Day boundaries are computed in the user's timezone; pure hour offsets are zone-independent.
  const today = datedPartsInTimeZone(now, timezone);
  const dayStart = (p: { year: number; month: number; day: number }) => zonedDayStartInstant(p.year, p.month, p.day, timezone);
  const end = new Date(now);
  let start = new Date(now);
  let resolution = "relative_window";
  if (window === "yesterday") {
    const y = addCalendarDays(today.year, today.month, today.day, -1);
    return { window, start: dayStart(y).toISOString(), end: dayStart(today).toISOString(), date_text: dateText || undefined, timezone, resolution: "yesterday" };
  } else if (window === "since_yesterday") {
    start = dayStart(addCalendarDays(today.year, today.month, today.day, -1));
  } else if (window === "last_48_hours") {
    start.setTime(now.getTime() - 48 * 60 * 60 * 1000);
  } else if (window === "last_24_hours") {
    start.setTime(now.getTime() - 24 * 60 * 60 * 1000);
  } else if (window === "last_7_days") {
    start.setTime(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  } else if (window === "last_30_days") {
    start.setTime(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  } else if (window === "this_week") {
    start = dayStart(addCalendarDays(today.year, today.month, today.day, -((today.weekday + 6) % 7)));
  } else if (window === "last_week") {
    const monday = previousMondayParts(now, timezone);
    return { window, start: dayStart(monday).toISOString(), end: dayStart(addCalendarDays(monday.year, monday.month, monday.day, 7)).toISOString(), date_text: dateText || undefined, timezone, resolution: "last_week" };
  } else {
    start = dayStart(today);
    resolution = "today";
  }
  return { window, start: start.toISOString(), end: end.toISOString(), date_text: dateText || undefined, timezone, resolution };
}

export function emailQuestionTypeFromArgs(args: Record<string, unknown>): EmailQuestionType {
  const supplied = typeof args.question_type === "string" ? args.question_type.trim() : "";
  if (["morning_brief", "attention_digest", "recent_unread", "today_digest", "since_yesterday"].includes(supplied)) return supplied as EmailQuestionType;
  const q = typeof args.question === "string" ? args.question.toLowerCase() : "";
  if (q.includes("since yesterday")) return "since_yesterday";
  if (q.includes("unread")) return "recent_unread";
  if (q.includes("pay attention") || q.includes("important") || q.includes("urgent")) return "attention_digest";
  if (q.includes("morning") || q.includes("brief")) return "morning_brief";
  return "today_digest";
}

export function emailMaxResultsFromArgs(args: Record<string, unknown>): number {
  const filters = emailFiltersFromArgs(args);
  const q = typeof args.question === "string" ? args.question : "";
  const minimum = filters.has_filters || /\b(last time|last email|emailed me|from)\b/i.test(q) ? 50 : 10;
  const raw = typeof args.max_results === "number" ? args.max_results : Number(args.max_results || minimum);
  if (!Number.isFinite(raw)) return minimum;
  return Math.max(minimum, Math.min(50, Math.floor(raw)));
}

export function emailFolderFromArgs(args: Record<string, unknown>): EmailFolderScope {
  const explicit = typeof args.folderScope === "string" ? args.folderScope.trim().toLowerCase() : (typeof args.folder === "string" ? args.folder.trim().toLowerCase() : "");
  if (explicit === "inbox") return "inbox";
  if (explicit === "mailbox_basic" || explicit === "mailbox" || explicit === "all_basic") return "mailbox_basic";
  const q = typeof args.question === "string" ? args.question.toLowerCase() : "";
  if (/\b(last time|last email|emailed me|my email|mailbox)\b/.test(q)) return "mailbox_basic";
  return "inbox";
}

export function safeSenderDisplay(value: unknown): string | null {
  const from = value && typeof value === "object" ? value as Record<string, unknown> : {};
  const emailAddress = from.emailAddress && typeof from.emailAddress === "object" ? from.emailAddress as Record<string, unknown> : {};
  const display = safeCalendarText(emailAddress.name || from.name || from.displayName, "");
  return display || null;
}

export function rawSenderAddress(value: unknown): string {
  const from = value && typeof value === "object" ? value as Record<string, unknown> : {};
  const emailAddress = from.emailAddress && typeof from.emailAddress === "object" ? from.emailAddress as Record<string, unknown> : {};
  return typeof emailAddress.address === "string" ? emailAddress.address.toLowerCase() : "";
}

export function safeSenderDomainLabel(value: unknown): string | null {
  const address = rawSenderAddress(value);
  const domain = address.includes("@") ? address.split("@").pop() || "" : "";
  if (!/^[a-z0-9.-]+\.[a-z]{2,}$/i.test(domain)) return null;
  return domain;
}

export function normalizeSearchText(value: unknown): string {
  return String(value || "").toLowerCase().replace(/<https?:\/\/[^|>]+\|([^>]+)>/g, "$1").replace(/[^a-z0-9.]+/g, " ").replace(/\s+/g, " ").trim();
}

export function expandEmailSearchAliases(terms: string[]): string[] {
  const out = new Set<string>();
  for (const term of terms) {
    if (!term) continue;
    out.add(term);
    if (term === "monday" || term === "monday.com" || term === "mondaycom") {
      out.add("monday");
      out.add("monday.com");
      out.add("mondaycom");
    }
    if (term === "isiah" || term === "isaiah") {
      out.add("isiah");
      out.add("isaiah");
    }
  }
  return [...out];
}

export function emailFiltersFromArgs(args: Record<string, unknown>): EmailQueryFilters {
  const senderOrDomain = typeof args.senderOrDomain === "string" ? args.senderOrDomain.trim() : "";
  const subjectContains = typeof args.subjectContains === "string" ? args.subjectContains.trim() : "";
  const keyword = typeof args.keyword === "string" ? args.keyword.trim() : "";
  const question = typeof args.question === "string" ? args.question.trim() : "";
  const explicitTerms = [senderOrDomain, subjectContains, keyword].map(normalizeSearchText).filter(Boolean);
  const questionTerms: string[] = [];
  const q = normalizeSearchText(question);
  if (q.includes("monday.com") || q.includes("monday com") || /\bmonday\b/.test(q)) questionTerms.push("monday.com");
  const quoted = [...question.matchAll(/"([^"]+)"/g)].map((match) => normalizeSearchText(match[1])).filter(Boolean);
  const terms = [...explicitTerms, ...questionTerms, ...quoted];
  const aliasTerms = expandEmailSearchAliases(terms);
  return { senderOrDomain: senderOrDomain || undefined, subjectContains: subjectContains || undefined, keyword: keyword || undefined, normalized_terms: terms, alias_terms: aliasTerms, has_filters: aliasTerms.length > 0 };
}

export function emailMessageMatchesFilters(message: EmailMetadataMessage, filters: EmailQueryFilters): boolean {
  if (!filters.has_filters) return true;
  const subject = normalizeSearchText(message.subject);
  const display = normalizeSearchText(safeSenderDisplay(message.from));
  const domain = normalizeSearchText(safeSenderDomainLabel(message.from));
  const compactDomain = domain.replace(/\s+/g, "");
  const haystack = [subject, display, domain, compactDomain].join(" ");
  return filters.alias_terms.some((term) => {
    const normalized = normalizeSearchText(term);
    const compact = normalized.replace(/\s+/g, "");
    return Boolean(normalized) && (haystack.includes(normalized) || (compact.length > 2 && haystack.includes(compact)));
  });
}

export function preferenceKeywordMatches(preferences: Record<string, unknown>, subject: string): boolean {
  const subjectRules = Array.isArray(preferences.subject_keyword_rules) ? preferences.subject_keyword_rules as Record<string, unknown>[] : [];
  return subjectRules.some((rule) => rule && typeof rule === "object" && rule.status === "active" && typeof rule.label === "string" && subject.includes(rule.label.toLowerCase()));
}

export function scoreEmailImportance(summary: Record<string, unknown>, preferences: Record<string, unknown>): { score: number; label: string; reasons: string[] } {
  const subject = typeof summary.subject === "string" ? summary.subject.toLowerCase() : "";
  const reasons: string[] = [];
  let score = 0;
  if (String(summary.importance || "").toLowerCase() === "high") { score += 3; reasons.push("Microsoft 365 marks it high importance"); }
  if (summary.is_unread === true) { score += 1; reasons.push("it is unread"); }
  const matched = importantEmailKeywords.filter((keyword) => subject.includes(keyword));
  if (matched.length) { score += 2; reasons.push("the subject has business-priority language"); }
  if (summary.has_attachments === true) { score += 1; reasons.push("it has an attachment flag"); }
  if (preferenceKeywordMatches(preferences, subject)) { score += 2; reasons.push("it matches one of your saved subject keyword preferences"); }
  const received = typeof summary.received_at === "string" ? new Date(summary.received_at).getTime() : NaN;
  if (Number.isFinite(received) && Date.now() - received <= 4 * 60 * 60 * 1000) { score += 1; reasons.push("it arrived recently"); }
  const label = score >= 4 ? "likely needs attention" : score >= 2 ? "may be worth attention" : "routine metadata signal";
  return { score, label, reasons: reasons.slice(0, 4) };
}

export function summarizeEmailMessage(message: EmailMetadataMessage, preferences: Record<string, unknown>): Record<string, unknown> {
  const id = typeof message.id === "string" ? message.id : JSON.stringify(message).slice(0, 120);
  const conversationId = typeof message.conversationId === "string" ? message.conversationId : "";
  const summary: Record<string, unknown> = {
    message_id_hash: sha256Hex(id).slice(0, 16),
    conversation_id_hash: conversationId ? sha256Hex(conversationId).slice(0, 16) : null,
    subject: safeCalendarText(message.subject, "No subject"),
    sender_display: safeSenderDisplay(message.from),
    sender_domain_label: safeSenderDomainLabel(message.from),
    received_at: typeof message.receivedDateTime === "string" ? message.receivedDateTime : null,
    importance: typeof message.importance === "string" ? safeCalendarText(message.importance, "normal") : "normal",
    is_unread: message.isRead === false,
    has_attachments: message.hasAttachments === true,
    folder_label: "inbox",
  };
  summary.importance_signal = scoreEmailImportance(summary, preferences);
  return summary;
}

export function appendEmailQueryAudit(event: Record<string, unknown>): void {
  mkdirSync(EMAIL_QUERY_AUDIT_PATH.split("/").slice(0, -1).join("/"), { recursive: true, mode: 0o700 });
  const metadata = redact({ timestamp_utc: isoNow(), source: "emclaw_m365_email_query", tokens_logged: false, raw_slack_ids_logged: false, raw_emails_logged: false, m365_content_accessed: false, body_read: false, bodyPreview_read: false, attachments_read: false, mailbox_write: false, ...event });
  appendFileSync(EMAIL_QUERY_AUDIT_PATH, JSON.stringify(metadata) + "\n", { mode: 0o600 });
}

export function formatEmailQueryAnswer(caller: Caller, args: Record<string, unknown>, messages: EmailMetadataMessage[], graphCalled: boolean, auditExtra: Record<string, unknown> = {}): ToolResult {
  const preferences = ensurePreferenceRecord(caller);
  const displayName = activePreferenceValue(preferences, "display_name_preference");
  const questionType = emailQuestionTypeFromArgs(args);
  const window = emailWindowFromArgs(args);
  const folderScope = emailFolderFromArgs(args);
  const filters = emailFiltersFromArgs(args);
  const matchedMessages = messages.filter((message) => emailMessageMatchesFilters(message, filters));
  const summaries = matchedMessages.map((message) => summarizeEmailMessage(message, preferences));
  const filtered = questionType === "recent_unread" ? summaries.filter((summary) => summary.is_unread === true) : summaries;
  const sorted = [...filtered].sort((a, b) => ((b.importance_signal as { score?: number }).score || 0) - ((a.importance_signal as { score?: number }).score || 0));
  const top = sorted.slice(0, Math.min(sorted.length, 5));
  const greeting = displayName ? safeCalendarText(displayName, "") + ", " : "";
  const rangeText = window.start.slice(0, 10) + " through " + window.end.slice(0, 10);
  let answer = greeting + "I found " + filtered.length + " email" + (filtered.length === 1 ? "" : "s") + " in " + (folderScope === "mailbox_basic" ? "your mailbox" : "your inbox") + " for " + rangeText + ".";
  if (filtered.length === 0) {
    const criteria = filters.has_filters ? filters.alias_terms.join(", ") : "no extra keyword filter";
    answer += " Nothing matched the basic details I checked. Range: " + rangeText + "; scope: " + folderScope + "; criteria: " + criteria + "; fetched: " + messages.length + ".";
  } else {
    const first = top[0];
    const signal = first.importance_signal as { label?: string; reasons?: string[] };
    answer += " The message most likely to matter is \"" + first.subject + "\" from " + (first.sender_display || "a sender") + "; it " + (signal.label || "has basic-detail signals") + ".";
    if (Array.isArray(signal.reasons) && signal.reasons.length) answer += " Why: " + signal.reasons.join("; ") + ".";
  }
  answer += " I only checked things like sender, subject, time, unread status, importance, and whether there is an attachment. I did not open or read any emails, previews, attachments, images, or links.";
  const safeFilterAudit = { has_filters: filters.has_filters, alias_terms: filters.alias_terms, sender_or_domain_present: Boolean(filters.senderOrDomain), subject_contains_present: Boolean(filters.subjectContains), keyword_present: Boolean(filters.keyword) };
  appendEmailQueryAudit({ caller_user_ref: caller.user_ref, operation: "m365_email_query", question_type: questionType, window: window.window, resolved_start: window.start, resolved_end: window.end, date_text: window.date_text, date_resolution: window.resolution, timezone: window.timezone, folder_scope: folderScope, fetched_count: messages.length, result_count: filtered.length, graph_called: graphCalled, graph_call_scope: graphCalled ? "mail_metadata_only" : "none", content_access: "metadata_only", body_read: false, bodyPreview_read: false, uniqueBody_read: false, attachments_read: false, mailbox_write: false, tokens_logged: false, raw_emails_logged: false, raw_slack_ids_logged: false, filters: safeFilterAudit, ...auditExtra });
  return okResult({ action: "m365_email_query", operation_id: questionType === "morning_brief" || questionType === "today_digest" ? "mail.metadata_digest" : "mail.question_answer", user_ref: caller.user_ref, window, folder_scope: folderScope, filters: safeFilterAudit, fetched_count: messages.length, question_type: questionType, answer_text: answer, messages: top, total_matching_messages: filtered.length, allowed_fields: emailQuerySelectFields, excluded_fields: ["body", "bodyPreview", "uniqueBody", "mimeContent", "internetMessageHeaders", "attachments", "attachment content", "images", "OCR", "raw email addresses", "recipients", "mailbox writes"], graph_called: graphCalled, graph_call_scope: graphCalled ? "mail_metadata_only" : "none", content_access: "metadata_only", email_body_read: false, body_preview_read: false, unique_body_read: false, attachments_read: false, headers_read: false, mime_read: false, mailbox_write: false, raw_email_addresses_exposed: false, workflow_created: false, ...auditExtra });
}

export function graphMailMessagesUrl(window: { start: string; end: string }, maxResults: number, folderScope: EmailFolderScope, nextLink?: string): string {
  if (nextLink) return nextLink;
  const params = new URLSearchParams();
  params.set("$select", emailQuerySelectFields.join(","));
  params.set("$orderby", "receivedDateTime desc");
  params.set("$top", String(maxResults));
  params.set("$filter", "receivedDateTime ge " + window.start + " and receivedDateTime lt " + window.end);
  const base = folderScope === "mailbox_basic" ? "https://graph.microsoft.com/v1.0/me/messages?" : "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?";
  return base + params.toString();
}

async function graphMailFetch(accessToken: string, window: { start: string; end: string }, maxResults: number, folderScope: EmailFolderScope, nextLink?: string): Promise<GraphFetchResult> {
  const fetchImpl = (globalThis as unknown as { fetch?: (input: string, init?: Record<string, unknown>) => Promise<{ ok: boolean; status: number; json: () => Promise<unknown> }> }).fetch;
  if (typeof fetchImpl !== "function") throw new Error("fetch_unavailable");
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 15000);
  try {
    const response = await fetchImpl(graphMailMessagesUrl(window, maxResults, folderScope, nextLink), { method: "GET", headers: { Authorization: "Bearer " + accessToken, Accept: "application/json" }, signal: controller.signal });
    const data = await response.json().catch(() => ({}));
    return { ok: response.ok === true, status: response.status, data };
  } finally {
    clearTimeout(timeout);
  }
}

export function mailMessagesFromGraphData(data: unknown): EmailMetadataMessage[] {
  const value = data && typeof data === "object" ? (data as { value?: unknown }).value : undefined;
  return Array.isArray(value) ? value.filter((item) => item && typeof item === "object") as EmailMetadataMessage[] : [];
}

export type GraphMailResult = { messages: EmailMetadataMessage[]; meta: Record<string, unknown> };

export function nextLinkFromGraphData(data: unknown): string | undefined {
  const value = data && typeof data === "object" ? (data as Record<string, unknown>)["@odata.nextLink"] : undefined;
  return typeof value === "string" && value.startsWith("https://graph.microsoft.com/") ? value : undefined;
}

async function graphMailFetchAll(accessToken: string, window: { start: string; end: string }, maxResults: number, folderScope: EmailFolderScope): Promise<{ messages: EmailMetadataMessage[]; status: number; page_count: number }> {
  const messages: EmailMetadataMessage[] = [];
  let nextLink: string | undefined;
  let status = 0;
  let pageCount = 0;
  do {
    const remaining = Math.max(1, maxResults - messages.length);
    const page = await graphMailFetch(accessToken, window, Math.min(50, remaining), folderScope, nextLink);
    status = page.status;
    if (page.ok !== true) return { messages, status, page_count: pageCount + 1 };
    messages.push(...mailMessagesFromGraphData(page.data));
    nextLink = nextLinkFromGraphData(page.data);
    pageCount += 1;
  } while (nextLink && messages.length < maxResults && pageCount < 4);
  return { messages: messages.slice(0, maxResults), status, page_count: pageCount };
}

async function fetchOwnMailMetadata(userRef: string, window: { start: string; end: string }, maxResults: number, folderScope: EmailFolderScope): Promise<GraphMailResult> {
  const load = loadTokenWrapperForGraph(userRef);
  const first = await graphMailFetchAll(accessTokenFromLoad(load), window, maxResults, folderScope);
  if (first.status >= 200 && first.status < 300) return { messages: first.messages, meta: { graph_call_count: first.page_count, initial_http_status: first.status, refresh_attempted: false, refresh_succeeded: false, token_refreshed: false, folder_scope: folderScope } };
  if (first.status !== 401) throw new Error("graph_mail_metadata_http_" + first.status);
  await refreshAccessToken(load);
  const second = await graphMailFetchAll(accessTokenFromLoad(load), window, maxResults, folderScope);
  if (second.status < 200 || second.status >= 300) throw new Error("graph_mail_metadata_http_" + second.status);
  return { messages: second.messages, meta: { graph_call_count: first.page_count + second.page_count, initial_http_status: first.status, final_http_status: second.status, refresh_attempted: true, refresh_succeeded: true, token_refreshed: true, folder_scope: folderScope } };
}

export async function m365EmailQueryForCaller(args: Record<string, unknown>, caller: Caller): Promise<ToolResult> {
  if (tokenFileExistsByStat(caller.user_ref) !== true) return errorResult("microsoft365_connection_not_ready_by_token_stat", "Microsoft 365 is not connected for this member by private token-file stat.");
  const tzResolved = await timezoneAwareArgs(args, caller);
  args = tzResolved.args;
  const tzMeta = { timezone_source: tzResolved.source, mailbox_timezone_status: tzResolved.status };
  const window = emailWindowFromArgs(args);
  const maxResults = emailMaxResultsFromArgs(args);
  const folderScope = emailFolderFromArgs(args);
  try {
    const result = await fetchOwnMailMetadata(caller.user_ref, window, maxResults, folderScope);
    return formatEmailQueryAnswer(caller, args, result.messages, true, { ...result.meta, ...tzMeta, user_timezone: window.timezone });
  } catch (err) {
    const errorCode = err instanceof Error ? err.message : String(err);
    const reconnectNeeded = errorCode === "refresh_token_missing" || errorCode === "refresh_token_expired_or_invalid";
    appendEmailQueryAudit({ caller_user_ref: caller.user_ref, operation: "m365_email_query", question_type: emailQuestionTypeFromArgs(args), window: window.window, result_count: 0, graph_called: true, graph_call_scope: "mail_metadata_only", content_access: "metadata_only", body_read: false, bodyPreview_read: false, uniqueBody_read: false, attachments_read: false, mailbox_write: false, tokens_logged: false, raw_emails_logged: false, raw_slack_ids_logged: false, refresh_attempted: errorCode !== "access_token_missing", refresh_succeeded: false, reconnect_required: reconnectNeeded, error_code: errorCode });
    return errorResult(reconnectNeeded ? "microsoft365_reconnect_required" : "m365_email_query_failed", reconnectNeeded ? "Microsoft 365 needs to be reconnected before I can check your basic email details." : errorCode, { action: "m365_email_query", user_ref: caller.user_ref, window, question_type: emailQuestionTypeFromArgs(args), allowed_fields: emailQuerySelectFields, excluded_fields: ["body", "bodyPreview", "uniqueBody", "mimeContent", "internetMessageHeaders", "attachments", "attachment content", "images", "OCR", "raw email addresses", "mailbox writes"], graph_called: true, graph_call_scope: "mail_metadata_only", content_access: "metadata_only", email_body_read: false, body_preview_read: false, unique_body_read: false, attachments_read: false, headers_read: false, mime_read: false, mailbox_write: false, workflow_created: false, refresh_succeeded: false, reconnect_required: reconnectNeeded, token_contents_logged_or_exposed: false });
  }
}

export function dryM365EmailQueryForCaller(caller: Caller, args: Record<string, unknown>, messages: EmailMetadataMessage[]): ToolResult {
  return formatEmailQueryAnswer(caller, args, messages, false);
}

