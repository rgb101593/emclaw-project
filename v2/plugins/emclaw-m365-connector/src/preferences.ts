import { existsSync } from "node:fs";

import { PREFERENCES_ROOT } from "./config.js";
import { ensureDir, isoNow, okResult, readJsonFile, redact, safeFileNamePart, writePrivateJsonAtomic } from "./util.js";
import type { Caller, ToolResult } from "./types.js";

const preferenceStringKeys = new Set(["display_name_preference", "tone_preference", "response_format_preference", "timezone_preference"]);
const preferenceObjectKeys = new Set(["email_digest_preferences", "email_alert_preferences", "calendar_alert_preferences"]);
const preferenceArrayKeys = new Set(["priority_senders_or_domains", "subject_keyword_rules", "calendar_keyword_rules"]);
const allPreferenceKeys = new Set([...preferenceStringKeys, ...preferenceObjectKeys, ...preferenceArrayKeys]);


export function preferenceDirFor(userRef: string): string {
  safeFileNamePart(userRef);
  return `${PREFERENCES_ROOT}/${userRef}`;
}

export function preferencePathFor(userRef: string): string {
  return `${preferenceDirFor(userRef)}/preferences.json`;
}

export function pendingPreferencePathFor(userRef: string): string {
  return `${preferenceDirFor(userRef)}/pending-preferences.json`;
}

export function defaultPreferenceRecord(userRef: string): Record<string, unknown> {
  const now = isoNow();
  const stringDefault = { value: null, status: "active", source: "default", created_at_utc: now, updated_at_utc: now, last_confirmed_at_utc: null };
  return {
    schema_version: "2026-05-28.m7_5_m8_0",
    user_ref: userRef,
    display_name_preference: { ...stringDefault },
    timezone_preference: { ...stringDefault },
    tone_preference: { ...stringDefault, value: "warm_concise_secretary" },
    response_format_preference: { ...stringDefault, value: "plain_language_concise" },
    email_digest_preferences: { status: "active", source: "default" },
    email_alert_preferences: { status: "active", source: "default" },
    calendar_alert_preferences: { status: "active", source: "default" },
    priority_senders_or_domains: [],
    subject_keyword_rules: [],
    calendar_keyword_rules: [],
    status: "active",
    source: "default",
    created_at_utc: now,
    updated_at_utc: now,
    last_confirmed_at_utc: null,
    audit_without_content: true,
    user_inspection_allowed: true,
    user_delete_allowed: true,
    admin_content_visibility: false,
    admin_metadata_visibility: true,
  };
}

export function ensurePreferenceRecord(caller: Caller): Record<string, unknown> {
  ensureDir(preferenceDirFor(caller.user_ref));
  const prefPath = preferencePathFor(caller.user_ref);
  if (existsSync(prefPath) === false) {
    writePrivateJsonAtomic(prefPath, defaultPreferenceRecord(caller.user_ref));
  }
  const parsed = readJsonFile(prefPath);
  if (!parsed || typeof parsed !== "object") throw new Error("invalid_preference_record");
  return parsed as Record<string, unknown>;
}

export function assertPreferenceKey(value: unknown): string {
  if (typeof value !== "string" || allPreferenceKeys.has(value) === false) throw new Error("unsupported_preference_key");
  return value;
}

export function safePreferenceValue(value: unknown): string {
  if (typeof value !== "string" || value.trim().length === 0 || value.length > 240) throw new Error("invalid_preference_value");
  const redacted = redact(value);
  if (typeof redacted !== "string" || redacted !== value) throw new Error("preference_value_contains_private_identifier");
  return value.trim();
}

export function preferenceSource(value: unknown, fallback = "explicit"): "explicit" | "confirmed_inferred" | "default" {
  if (value === "explicit" || value === "confirmed_inferred" || value === "default") return value;
  return fallback as "explicit";
}

export function showMyPreferencesForCaller(caller: Caller): ToolResult {
  const preferences = ensurePreferenceRecord(caller);
  return okResult({ action: "show_my_preferences", user_ref: caller.user_ref, preferences, admin_content_visibility: false, user_inspection_allowed: true, token_contents_read: false, m365_content_accessed: false });
}

export function setMyPreferenceForCaller(args: Record<string, unknown>, caller: Caller): ToolResult {
  const key = assertPreferenceKey(args.key);
  const value = safePreferenceValue(args.value);
  const source = preferenceSource(args.source, "explicit");
  if (source === "default") throw new Error("default_source_not_allowed_for_member_set");
  const preferences = ensurePreferenceRecord(caller);
  const now = isoNow();
  if (preferenceStringKeys.has(key)) {
    preferences[key] = { value, status: "active", source, created_at_utc: (preferences[key] as Record<string, unknown> | undefined)?.created_at_utc || now, updated_at_utc: now, last_confirmed_at_utc: now };
  } else if (preferenceObjectKeys.has(key)) {
    const current = preferences[key] && typeof preferences[key] === "object" ? preferences[key] as Record<string, unknown> : {};
    preferences[key] = { ...current, value, status: "active", source, updated_at_utc: now, last_confirmed_at_utc: now };
  } else if (preferenceArrayKeys.has(key)) {
    const arr = Array.isArray(preferences[key]) ? preferences[key] as unknown[] : [];
    arr.push({ label: value, redacted_or_safe: true, status: "active", source, created_at_utc: now, updated_at_utc: now, last_confirmed_at_utc: now });
    preferences[key] = arr;
  }
  preferences.status = "active";
  preferences.updated_at_utc = now;
  preferences.last_confirmed_at_utc = now;
  preferences.source = source;
  writePrivateJsonAtomic(preferencePathFor(caller.user_ref), preferences);
  return okResult({ action: "set_my_preference", user_ref: caller.user_ref, key, status: "active", source, preference_content_visible_to_admin: false });
}

export function forgetMyPreferenceForCaller(args: Record<string, unknown>, caller: Caller): ToolResult {
  const key = assertPreferenceKey(args.key);
  const label = typeof args.value === "string" && args.value.trim() ? args.value.trim() : "";
  const preferences = ensurePreferenceRecord(caller);
  const now = isoNow();
  if (preferenceStringKeys.has(key)) {
    preferences[key] = { value: null, status: "forgotten", source: "explicit", created_at_utc: (preferences[key] as Record<string, unknown> | undefined)?.created_at_utc || now, updated_at_utc: now, last_confirmed_at_utc: now };
  } else if (preferenceObjectKeys.has(key)) {
    preferences[key] = { status: "forgotten", source: "explicit", updated_at_utc: now, last_confirmed_at_utc: now };
  } else if (preferenceArrayKeys.has(key)) {
    const arr = Array.isArray(preferences[key]) ? preferences[key] as Record<string, unknown>[] : [];
    preferences[key] = arr.map((item) => (!label || item.label === label) ? { ...item, status: "forgotten", updated_at_utc: now, last_confirmed_at_utc: now } : item);
  }
  preferences.updated_at_utc = now;
  preferences.last_confirmed_at_utc = now;
  writePrivateJsonAtomic(preferencePathFor(caller.user_ref), preferences);
  return okResult({ action: "forget_my_preference", user_ref: caller.user_ref, key, status: "forgotten", preference_content_visible_to_admin: false });
}

export function suggestMyPreferenceForCaller(args: Record<string, unknown>, caller: Caller): ToolResult {
  const key = assertPreferenceKey(args.key);
  const value = safePreferenceValue(args.value);
  ensureDir(preferenceDirFor(caller.user_ref));
  const pendingPath = pendingPreferencePathFor(caller.user_ref);
  const pending = existsSync(pendingPath) ? readJsonFile(pendingPath) : { schema_version: "2026-05-28.m7_6_pending", user_ref: caller.user_ref, pending: [] };
  const record = pending && typeof pending === "object" ? pending as Record<string, unknown> : { schema_version: "2026-05-28.m7_6_pending", user_ref: caller.user_ref, pending: [] };
  const arr = Array.isArray(record.pending) ? record.pending as unknown[] : [];
  arr.push({ key, value, source: "inferred", status: "needs_confirmation", created_at_utc: isoNow(), active: false });
  record.pending = arr;
  record.user_ref = caller.user_ref;
  writePrivateJsonAtomic(pendingPath, record);
  return okResult({ action: "suggest_my_preference", user_ref: caller.user_ref, key, status: "needs_confirmation", active: false, requires_member_confirmation: true, preference_content_visible_to_admin: false });
}
