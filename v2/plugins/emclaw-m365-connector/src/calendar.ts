import { appendFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";

import {
  APPROVED_SAFE_DISPLAY_NAMES_BY_REF, CALENDAR_QUERY_AUDIT_PATH,
  M365_CLIENT_CONFIG_PATH, PER_MEMBER_AGENT_REFS, PRIVATE_USER_DIRECTORY_PATH,
  ROSTER_PATH, TEAM_CALENDAR_AUDIT_PATH, TOKEN_DIR,
} from "./config.js";
import { ensurePreferenceRecord, preferencePathFor } from "./preferences.js";
import { loadRoster, requireSelfCapable, sha256Hex } from "./identity.js";
import {
  ensureDir, errorResult, isoNow, okResult, readJsonFile, redact,
  tokenFileExistsByStat, tokenFileNameFor, writePrivateJsonAtomic,
} from "./util.js";
import type {
  CalendarMetadataEvent, CalendarQueryWindow, CalendarQuestionType, Caller,
  RosterUser, TeamAvailabilityIntent, TeamScheduleTarget, ToolResult,
} from "./types.js";

export const calendarQuerySelectFields = ["id", "subject", "organizer", "start", "end", "showAs", "responseStatus", "attendees", "location", "isOnlineMeeting", "recurrence"];
export const importantCalendarKeywords = ["important", "client", "customer", "prospect", "exec", "board", "interview", "urgent", "deadline", "review", "legal", "finance", "renewal", "contract"];
export const joinUrlPattern = /https?:\/\/\S+|\b(?:zoom|teams|meet\.google|webex)\b/i;

export function activePreferenceValue(preferences: Record<string, unknown>, key: string): string | null {
  const item = preferences[key];
  if (!item || typeof item !== "object" || Array.isArray(item)) return null;
  const record = item as Record<string, unknown>;
  if (record.status !== "active" || typeof record.value !== "string" || record.value.trim() === "") return null;
  return record.value.trim();
}

export function safeCalendarText(value: unknown, fallback = "Untitled"): string {
  if (typeof value !== "string" || value.trim() === "") return fallback;
  const stripped = value.replace(joinUrlPattern, "[meeting link omitted]").trim();
  const redacted = redact(stripped);
  return typeof redacted === "string" && redacted.trim() ? redacted.trim().slice(0, 180) : fallback;
}

export function safeOrganizerDisplay(value: unknown): string | null {
  const organizer = value && typeof value === "object" ? value as Record<string, unknown> : {};
  const emailAddress = organizer.emailAddress && typeof organizer.emailAddress === "object" ? organizer.emailAddress as Record<string, unknown> : {};
  const display = safeCalendarText(emailAddress.name || organizer.name || organizer.displayName, "");
  return display || null;
}

export function eventDateValue(value: unknown): string | null {
  const record = value && typeof value === "object" ? value as Record<string, unknown> : {};
  return typeof record.dateTime === "string" ? record.dateTime : null;
}

export function attendeeCounts(value: unknown): { attendee_count: number; external_attendee_count: number } {
  if (!Array.isArray(value)) return { attendee_count: 0, external_attendee_count: 0 };
  let external = 0;
  for (const item of value) {
    const record = item && typeof item === "object" ? item as Record<string, unknown> : {};
    const emailAddress = record.emailAddress && typeof record.emailAddress === "object" ? record.emailAddress as Record<string, unknown> : {};
    const address = typeof emailAddress.address === "string" ? emailAddress.address.toLowerCase() : "";
    if (address && !address.endsWith("@acme.com")) external += 1;
  }
  return { attendee_count: value.length, external_attendee_count: external };
}

export function internalAttendeeCount(value: unknown): number {
  if (!Array.isArray(value)) return 0;
  let internal = 0;
  for (const item of value) {
    const record = item && typeof item === "object" ? item as Record<string, unknown> : {};
    const emailAddress = record.emailAddress && typeof record.emailAddress === "object" ? record.emailAddress as Record<string, unknown> : {};
    const address = typeof emailAddress.address === "string" ? emailAddress.address.toLowerCase() : "";
    if (address.endsWith("@acme.com")) internal += 1;
  }
  return internal;
}

export function hasSafeLocation(value: unknown): boolean {
  const location = value && typeof value === "object" ? value as Record<string, unknown> : {};
  const label = typeof location.displayName === "string" ? location.displayName : "";
  return Boolean(label.trim()) && !joinUrlPattern.test(label);
}

export function normalizeTimezoneName(value: unknown): string {
  const raw = typeof value === "string" ? value.trim() : "";
  const lower = raw.toLowerCase();
  if (!raw || lower === "et" || lower === "eastern" || lower === "eastern time" || lower === "est" || lower === "edt" || raw === "Eastern Standard Time") return "America/New_York";
  return raw;
}

export function graphWindowsTimezoneFor(value: string): string {
  const normalized = normalizeTimezoneName(value);
  if (normalized === "America/New_York") return "Eastern Standard Time";
  return normalized;
}

export function timezoneOffsetMinutes(instant: Date, timeZone: string): number {
  const parts = new Intl.DateTimeFormat("en-US", { timeZone, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23" }).formatToParts(instant);
  const get = (type: string) => Number(parts.find((part) => part.type === type)?.value || 0);
  const asUtc = Date.UTC(get("year"), get("month") - 1, get("day"), get("hour"), get("minute"), get("second"));
  return Math.round((asUtc - instant.getTime()) / 60000);
}

export function offsetText(minutes: number): string {
  const sign = minutes >= 0 ? "+" : "-";
  const abs = Math.abs(minutes);
  return `${sign}${String(Math.floor(abs / 60)).padStart(2, "0")}:${String(abs % 60).padStart(2, "0")}`;
}

export function isoWithTimezoneOffset(instantIso: string, timeZone: string): string {
  const instant = new Date(instantIso);
  const offset = timezoneOffsetMinutes(instant, timeZone);
  const parts = new Intl.DateTimeFormat("en-CA", { timeZone, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23" }).formatToParts(instant);
  const get = (type: string) => parts.find((part) => part.type === type)?.value || "00";
  return `${get("year")}-${get("month")}-${get("day")}T${get("hour")}:${get("minute")}:${get("second")}${offsetText(offset)}`;
}

export function instantFromZonedDateTime(dateTime: string, timeZone: unknown): Date {
  const normalizedTz = normalizeTimezoneName(timeZone);
  if (/Z$|[+-]\d{2}:?\d{2}$/.test(dateTime)) return new Date(dateTime);
  if (normalizedTz.toLowerCase() === "utc") return new Date(`${dateTime.replace(/\.\d+$/, "")}Z`);
  const match = dateTime.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?/);
  if (!match) return new Date(dateTime);
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = Number(match[4]);
  const minute = Number(match[5]);
  const second = Number(match[6] || 0);
  let guess = new Date(Date.UTC(year, month - 1, day, hour, minute, second));
  const offset = timezoneOffsetMinutes(guess, normalizedTz);
  guess = new Date(Date.UTC(year, month - 1, day, hour, minute, second) - offset * 60000);
  const correctedOffset = timezoneOffsetMinutes(guess, normalizedTz);
  if (correctedOffset !== offset) guess = new Date(Date.UTC(year, month - 1, day, hour, minute, second) - correctedOffset * 60000);
  return guess;
}

export function formatDisplayTime(instant: Date, timeZone: string): string {
  return new Intl.DateTimeFormat("en-US", { timeZone, hour: "numeric", minute: "2-digit", hour12: true }).format(instant);
}

// Microsoft Graph mailboxSettings.timeZone is usually a Windows time-zone name; Intl needs IANA.
// Covers the standard Windows zones (one representative IANA id each). Unknown values fall back.
export const windowsToIanaTimeZone: Record<string, string> = {
  "Dateline Standard Time": "Etc/GMT+12", "UTC-11": "Etc/GMT+11", "Aleutian Standard Time": "America/Adak", "Hawaiian Standard Time": "Pacific/Honolulu", "Marquesas Standard Time": "Pacific/Marquesas", "Alaskan Standard Time": "America/Anchorage", "UTC-09": "Etc/GMT+9", "Pacific Standard Time (Mexico)": "America/Tijuana", "UTC-08": "Etc/GMT+8", "Pacific Standard Time": "America/Los_Angeles", "US Mountain Standard Time": "America/Phoenix", "Mountain Standard Time (Mexico)": "America/Mazatlan", "Mountain Standard Time": "America/Denver", "Central America Standard Time": "America/Guatemala", "Central Standard Time": "America/Chicago", "Easter Island Standard Time": "Pacific/Easter", "Central Standard Time (Mexico)": "America/Mexico_City", "Canada Central Standard Time": "America/Regina", "SA Pacific Standard Time": "America/Bogota", "Eastern Standard Time (Mexico)": "America/Cancun", "Eastern Standard Time": "America/New_York", "Haiti Standard Time": "America/Port-au-Prince", "Cuba Standard Time": "America/Havana", "US Eastern Standard Time": "America/Indianapolis", "Turks And Caicos Standard Time": "America/Grand_Turk", "Paraguay Standard Time": "America/Asuncion", "Atlantic Standard Time": "America/Halifax", "Venezuela Standard Time": "America/Caracas", "Central Brazilian Standard Time": "America/Cuiaba", "SA Western Standard Time": "America/La_Paz", "Pacific SA Standard Time": "America/Santiago", "Newfoundland Standard Time": "America/St_Johns", "Tocantins Standard Time": "America/Araguaina", "E. South America Standard Time": "America/Sao_Paulo", "SA Eastern Standard Time": "America/Cayenne", "Argentina Standard Time": "America/Buenos_Aires", "Greenland Standard Time": "America/Godthab", "Montevideo Standard Time": "America/Montevideo", "Magallanes Standard Time": "America/Punta_Arenas", "Saint Pierre Standard Time": "America/Miquelon", "Bahia Standard Time": "America/Bahia", "UTC-02": "Etc/GMT+2", "Azores Standard Time": "Atlantic/Azores", "Cape Verde Standard Time": "Atlantic/Cape_Verde", "UTC": "Etc/UTC", "GMT Standard Time": "Europe/London", "Greenwich Standard Time": "Atlantic/Reykjavik", "Sao Tome Standard Time": "Africa/Sao_Tome", "Morocco Standard Time": "Africa/Casablanca", "W. Europe Standard Time": "Europe/Berlin", "Central Europe Standard Time": "Europe/Budapest", "Romance Standard Time": "Europe/Paris", "Central European Standard Time": "Europe/Warsaw", "W. Central Africa Standard Time": "Africa/Lagos", "Jordan Standard Time": "Asia/Amman", "GTB Standard Time": "Europe/Bucharest", "Middle East Standard Time": "Asia/Beirut", "Egypt Standard Time": "Africa/Cairo", "E. Europe Standard Time": "Europe/Chisinau", "Syria Standard Time": "Asia/Damascus", "West Bank Standard Time": "Asia/Hebron", "South Africa Standard Time": "Africa/Johannesburg", "FLE Standard Time": "Europe/Kiev", "Israel Standard Time": "Asia/Jerusalem", "Kaliningrad Standard Time": "Europe/Kaliningrad", "Sudan Standard Time": "Africa/Khartoum", "Libya Standard Time": "Africa/Tripoli", "Namibia Standard Time": "Africa/Windhoek", "Arabic Standard Time": "Asia/Baghdad", "Turkey Standard Time": "Europe/Istanbul", "Arab Standard Time": "Asia/Riyadh", "Belarus Standard Time": "Europe/Minsk", "Russian Standard Time": "Europe/Moscow", "E. Africa Standard Time": "Africa/Nairobi", "Iran Standard Time": "Asia/Tehran", "Arabian Standard Time": "Asia/Dubai", "Astrakhan Standard Time": "Europe/Astrakhan", "Azerbaijan Standard Time": "Asia/Baku", "Russia Time Zone 3": "Europe/Samara", "Mauritius Standard Time": "Indian/Mauritius", "Saratov Standard Time": "Europe/Saratov", "Georgian Standard Time": "Asia/Tbilisi", "Volgograd Standard Time": "Europe/Volgograd", "Caucasus Standard Time": "Asia/Yerevan", "Afghanistan Standard Time": "Asia/Kabul", "West Asia Standard Time": "Asia/Tashkent", "Ekaterinburg Standard Time": "Asia/Yekaterinburg", "Pakistan Standard Time": "Asia/Karachi", "India Standard Time": "Asia/Kolkata", "Sri Lanka Standard Time": "Asia/Colombo", "Nepal Standard Time": "Asia/Kathmandu", "Central Asia Standard Time": "Asia/Almaty", "Bangladesh Standard Time": "Asia/Dhaka", "Omsk Standard Time": "Asia/Omsk", "Myanmar Standard Time": "Asia/Yangon", "SE Asia Standard Time": "Asia/Bangkok", "Altai Standard Time": "Asia/Barnaul", "W. Mongolia Standard Time": "Asia/Hovd", "North Asia Standard Time": "Asia/Krasnoyarsk", "N. Central Asia Standard Time": "Asia/Novosibirsk", "Tomsk Standard Time": "Asia/Tomsk", "China Standard Time": "Asia/Shanghai", "North Asia East Standard Time": "Asia/Irkutsk", "Singapore Standard Time": "Asia/Singapore", "W. Australia Standard Time": "Australia/Perth", "Taipei Standard Time": "Asia/Taipei", "Ulaanbaatar Standard Time": "Asia/Ulaanbaatar", "Aus Central W. Standard Time": "Australia/Eucla", "Transbaikal Standard Time": "Asia/Chita", "Tokyo Standard Time": "Asia/Tokyo", "North Korea Standard Time": "Asia/Pyongyang", "Korea Standard Time": "Asia/Seoul", "Yakutsk Standard Time": "Asia/Yakutsk", "Cen. Australia Standard Time": "Australia/Adelaide", "AUS Central Standard Time": "Australia/Darwin", "E. Australia Standard Time": "Australia/Brisbane", "AUS Eastern Standard Time": "Australia/Sydney", "West Pacific Standard Time": "Pacific/Port_Moresby", "Tasmania Standard Time": "Australia/Hobart", "Vladivostok Standard Time": "Asia/Vladivostok", "Lord Howe Standard Time": "Australia/Lord_Howe", "Bougainville Standard Time": "Pacific/Bougainville", "Russia Time Zone 10": "Asia/Srednekolymsk", "Magadan Standard Time": "Asia/Magadan", "Norfolk Standard Time": "Pacific/Norfolk", "Sakhalin Standard Time": "Asia/Sakhalin", "Central Pacific Standard Time": "Pacific/Guadalcanal", "Russia Time Zone 11": "Asia/Kamchatka", "New Zealand Standard Time": "Pacific/Auckland", "UTC+12": "Etc/GMT-12", "Fiji Standard Time": "Pacific/Fiji", "Chatham Islands Standard Time": "Pacific/Chatham", "UTC+13": "Etc/GMT-13", "Tonga Standard Time": "Pacific/Tongatapu", "Samoa Standard Time": "Pacific/Apia", "Line Islands Standard Time": "Pacific/Kiritimati",
};

export function isUsableIanaTimeZone(value: string): boolean {
  try { new Intl.DateTimeFormat("en-US", { timeZone: value }); return true; } catch { return false; }
}

// Resolve a Graph mailboxSettings.timeZone string (Windows or IANA) to an IANA id Intl can use.
export function ianaFromGraphTimeZone(value: unknown): string | null {
  const raw = typeof value === "string" ? value.trim() : "";
  if (!raw) return null;
  if (windowsToIanaTimeZone[raw]) return windowsToIanaTimeZone[raw];
  const normalized = normalizeTimezoneName(raw);
  if (windowsToIanaTimeZone[normalized]) return windowsToIanaTimeZone[normalized];
  if (isUsableIanaTimeZone(normalized)) return normalized;
  return null;
}

// Accepts a user-supplied timezone: an IANA id, a Windows name, or a UTC/GMT offset like
// "GMT+8" / "UTC+8" / "+8" / "+08:00". Returns a usable IANA id (fixed-offset Etc/GMT zone for
// numeric offsets — note the IANA sign convention is inverted) or null.
export function userTimeZoneToIana(value: unknown): string | null {
  const raw = typeof value === "string" ? value.trim() : "";
  if (!raw) return null;
  const offset = raw.replace(/\s+/g, "").match(/^(?:gmt|utc)?([+-])(\d{1,2})(?::?(\d{2}))?$/i);
  if (offset) {
    const sign = offset[1] === "-" ? "+" : "-"; // Etc/GMT offsets are inverted
    const hours = Number(offset[2]);
    const minutes = Number(offset[3] || "0");
    if (hours >= 0 && hours <= 14 && minutes === 0) {
      const candidate = `Etc/GMT${sign}${hours}`;
      if (isUsableIanaTimeZone(candidate)) return candidate;
    }
  }
  return ianaFromGraphTimeZone(raw);
}

// Calendar-day parts (year/month/day/weekday 0=Sun) for an instant, evaluated in a given zone.
export function datedPartsInTimeZone(instant: Date, timeZone: string): { year: number; month: number; day: number; weekday: number } {
  const tz = normalizeTimezoneName(timeZone);
  const parts = new Intl.DateTimeFormat("en-US", { timeZone: tz, year: "numeric", month: "2-digit", day: "2-digit", weekday: "short", hourCycle: "h23" }).formatToParts(instant);
  const get = (type: string) => parts.find((part) => part.type === type)?.value || "";
  const wmap: Record<string, number> = { Sun: 0, Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6 };
  return { year: Number(get("year")), month: Number(get("month")), day: Number(get("day")), weekday: wmap[get("weekday")] ?? 0 };
}

// UTC instant for local midnight of (year,month,day) in the given zone.
export function zonedDayStartInstant(year: number, month: number, day: number, timeZone: string): Date {
  const pad = (n: number) => String(n).padStart(2, "0");
  return instantFromZonedDateTime(`${year}-${pad(month)}-${pad(day)}T00:00:00`, normalizeTimezoneName(timeZone));
}

// Add n calendar days to a (year,month,day) triple, DST-safe (pure date arithmetic).
export function addCalendarDays(year: number, month: number, day: number, n: number): { year: number; month: number; day: number } {
  const d = new Date(Date.UTC(year, month - 1, day));
  d.setUTCDate(d.getUTCDate() + n);
  return { year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate() };
}

// Preserves the original previousMonday() semantics (this week's Monday, or 7 days back when today is Monday), in the user's zone.
export function previousMondayParts(now: Date, timeZone: string): { year: number; month: number; day: number } {
  const t = datedPartsInTimeZone(now, timeZone);
  const sinceMonday = (t.weekday + 6) % 7;
  return addCalendarDays(t.year, t.month, t.day, sinceMonday === 0 ? -7 : -sinceMonday);
}

export function eventInstant(value: unknown): Date | null {
  const record = value && typeof value === "object" ? value as Record<string, unknown> : {};
  return typeof record.dateTime === "string" ? instantFromZonedDateTime(record.dateTime, record.timeZone) : null;
}

export function eventDisplayValue(value: unknown, timeZone: string): string | null {
  const instant = eventInstant(value);
  return instant ? formatDisplayTime(instant, timeZone) : null;
}

const monthNames: Record<string, number> = {
  january: 1, jan: 1,
  february: 2, feb: 2,
  march: 3, mar: 3,
  april: 4, apr: 4,
  may: 5,
  june: 6, jun: 6,
  july: 7, jul: 7,
  august: 8, aug: 8,
  september: 9, sep: 9, sept: 9,
  october: 10, oct: 10,
  november: 11, nov: 11,
  december: 12, dec: 12,
};

const weekdayNames: Record<string, number> = {
  sunday: 0, sun: 0,
  monday: 1, mon: 1,
  tuesday: 2, tue: 2, tues: 2,
  wednesday: 3, wed: 3,
  thursday: 4, thu: 4, thur: 4, thurs: 4,
  friday: 5, fri: 5,
  saturday: 6, sat: 6,
};

function inferYearForMonthDay(month: number, day: number, now: Date, timeZone: string): number {
  const thisYear = datedPartsInTimeZone(now, timeZone).year;
  const candidate = zonedDayStartInstant(thisYear, month, day, timeZone);
  if (candidate.getTime() - now.getTime() > 31 * 24 * 60 * 60 * 1000) return thisYear - 1;
  return thisYear;
}

// Parses natural date phrases into [start,end) instants, with ALL day boundaries anchored in the
// user's timezone (timeZone). Shared by calendar and email window resolution.
export function parseExplicitDateRange(text: string, now: Date, timeZone: string): { start: Date; end: Date; resolution: string } | null {
  const tz = normalizeTimezoneName(timeZone);
  const dayStart = (p: { year: number; month: number; day: number }) => zonedDayStartInstant(p.year, p.month, p.day, tz);
  const normalized = text.toLowerCase().replace(/<https?:\/\/[^|>]+\|([^>]+)>/g, "$1");
  const monthDay = normalized.match(/\b(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t|tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\.?\s+([0-3]?\d)(?:st|nd|rd|th)?(?:,\s*(20\d{2}))?\b/);
  if (monthDay) {
    const month = monthNames[monthDay[1].replace(".", "")] || monthNames[monthDay[1]];
    const day = Number(monthDay[2]);
    const year = monthDay[3] ? Number(monthDay[3]) : inferYearForMonthDay(month, day, now, tz);
    const start = { year, month, day };
    return { start: dayStart(start), end: dayStart(addCalendarDays(year, month, day, 1)), resolution: "absolute_day" };
  }
  const weekOf = normalized.match(/\bweek of\s+(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t|tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\.?\s+([0-3]?\d)(?:st|nd|rd|th)?(?:,\s*(20\d{2}))?\b/);
  if (weekOf) {
    const month = monthNames[weekOf[1].replace(".", "")] || monthNames[weekOf[1]];
    const day = Number(weekOf[2]);
    const year = weekOf[3] ? Number(weekOf[3]) : inferYearForMonthDay(month, day, now, tz);
    const anchorWeekday = datedPartsInTimeZone(zonedDayStartInstant(year, month, day, tz), tz).weekday;
    const monday = addCalendarDays(year, month, day, -((anchorWeekday + 6) % 7));
    return { start: dayStart(monday), end: dayStart(addCalendarDays(monday.year, monday.month, monday.day, 7)), resolution: "week_of" };
  }
  if (normalized.includes("last monday-friday") || normalized.includes("last monday to friday") || normalized.includes("previous work week")) {
    const monday = previousMondayParts(now, tz);
    return { start: dayStart(monday), end: dayStart(addCalendarDays(monday.year, monday.month, monday.day, 5)), resolution: "previous_business_week" };
  }
  if (normalized.includes("from monday to friday")) {
    const monday = previousMondayParts(now, tz);
    return { start: dayStart(monday), end: dayStart(addCalendarDays(monday.year, monday.month, monday.day, 5)), resolution: "monday_to_friday" };
  }
  if (normalized.includes("last week")) {
    const monday = previousMondayParts(now, tz);
    return { start: dayStart(monday), end: dayStart(addCalendarDays(monday.year, monday.month, monday.day, 7)), resolution: "last_week" };
  }
  const weekday = normalized.match(/\b(this|next|last|previous|prev)?\s*(sun(?:day)?|mon(?:day)?|tue(?:s|sday|day)?|wed(?:nesday)?|thu(?:r|rs|rsday|rday)?|fri(?:day)?|sat(?:urday)?)\b/);
  if (weekday) {
    const qualifier = weekday[1] || "this";
    const target = weekdayNames[weekday[2]];
    const today = datedPartsInTimeZone(now, tz);
    if (qualifier === "last" || qualifier === "previous" || qualifier === "prev") {
      let daysSince = (today.weekday - target + 7) % 7;
      if (daysSince === 0) daysSince = 7;
      const start = addCalendarDays(today.year, today.month, today.day, -daysSince);
      return { start: dayStart(start), end: dayStart(addCalendarDays(start.year, start.month, start.day, 1)), resolution: `last_${weekday[2]}` };
    }
    let daysUntil = (target - today.weekday + 7) % 7;
    if (qualifier === "next" && daysUntil === 0) daysUntil = 7;
    const start = addCalendarDays(today.year, today.month, today.day, daysUntil);
    return { start: dayStart(start), end: dayStart(addCalendarDays(start.year, start.month, start.day, 1)), resolution: `${qualifier}_${weekday[2]}` };
  }
  return null;
}

export function calendarWindowFromArgs(args: Record<string, unknown>): { window: CalendarQueryWindow; timezone: string; start: string; end: string; start_with_offset: string; end_with_offset: string; timezone_source: string; date_text?: string; resolution: string } {
  const timezone = normalizeTimezoneName(typeof args.timezone === "string" && args.timezone.trim() ? args.timezone.trim() : "America/New_York");
  const timezone_source = typeof args.timezone === "string" && args.timezone.trim() ? "explicit_request" : "default";
  const now = new Date();
  const dateText = typeof args.dateText === "string" ? args.dateText.trim() : "";
  const question = typeof args.question === "string" ? args.question.trim() : "";
  const rawStart = typeof args.startDateTime === "string" ? args.startDateTime.trim() : "";
  const rawEnd = typeof args.endDateTime === "string" ? args.endDateTime.trim() : "";
  if (rawStart && rawEnd) {
    const startIso = new Date(rawStart).toISOString();
    const endIso = new Date(rawEnd).toISOString();
    return { window: "explicit_range", timezone, start: startIso, end: endIso, start_with_offset: isoWithTimezoneOffset(startIso, timezone), end_with_offset: isoWithTimezoneOffset(endIso, timezone), timezone_source, date_text: dateText || undefined, resolution: "explicit_start_end" };
  }
  const parsed = parseExplicitDateRange([dateText, question].filter(Boolean).join(" "), now, timezone);
  if (parsed) {
    const startIso = parsed.start.toISOString();
    const endIso = parsed.end.toISOString();
    return { window: "explicit_range", timezone, start: startIso, end: endIso, start_with_offset: isoWithTimezoneOffset(startIso, timezone), end_with_offset: isoWithTimezoneOffset(endIso, timezone), timezone_source, date_text: dateText || question || undefined, resolution: parsed.resolution };
  }
  const defaultWindow = questionTypeFromArgs(args) === "next_meeting" ? "next_24_hours" : "today";
  const raw = typeof args.window === "string" ? args.window.trim().toLowerCase() : defaultWindow;
  const window = (["today", "tomorrow", "this_afternoon", "next_24_hours", "yesterday", "last_7_days", "this_week", "last_week"].includes(raw) ? raw : "today") as CalendarQueryWindow;
  // All day boundaries below are computed in the user's timezone, not the server's.
  const today = datedPartsInTimeZone(now, timezone);
  const dayStart = (p: { year: number; month: number; day: number }) => zonedDayStartInstant(p.year, p.month, p.day, timezone);
  const pad = (n: number) => String(n).padStart(2, "0");
  let startInstant: Date;
  let endInstant: Date;
  let resolution = "relative_window";
  if (window === "tomorrow") {
    const t1 = addCalendarDays(today.year, today.month, today.day, 1);
    startInstant = dayStart(t1); endInstant = dayStart(addCalendarDays(t1.year, t1.month, t1.day, 1));
    resolution = "tomorrow";
  } else if (window === "yesterday") {
    const y = addCalendarDays(today.year, today.month, today.day, -1);
    startInstant = dayStart(y); endInstant = dayStart(today);
    resolution = "yesterday";
  } else if (window === "this_afternoon") {
    startInstant = instantFromZonedDateTime(`${today.year}-${pad(today.month)}-${pad(today.day)}T12:00:00`, timezone);
    endInstant = instantFromZonedDateTime(`${today.year}-${pad(today.month)}-${pad(today.day)}T18:00:00`, timezone);
    resolution = "this_afternoon";
  } else if (window === "next_24_hours") {
    startInstant = now; endInstant = new Date(now.getTime() + 24 * 60 * 60 * 1000);
    resolution = "next_24_hours";
  } else if (window === "last_7_days") {
    startInstant = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000); endInstant = now;
  } else if (window === "this_week") {
    const monday = addCalendarDays(today.year, today.month, today.day, -((today.weekday + 6) % 7));
    startInstant = dayStart(monday); endInstant = now;
    resolution = "this_week";
  } else if (window === "last_week") {
    const monday = previousMondayParts(now, timezone);
    startInstant = dayStart(monday); endInstant = dayStart(addCalendarDays(monday.year, monday.month, monday.day, 7));
    resolution = "last_week";
  } else {
    startInstant = dayStart(today); endInstant = dayStart(addCalendarDays(today.year, today.month, today.day, 1));
    resolution = "today";
  }
  const startIso = startInstant.toISOString();
  const endIso = endInstant.toISOString();
  return { window, timezone, start: startIso, end: endIso, start_with_offset: isoWithTimezoneOffset(startIso, timezone), end_with_offset: isoWithTimezoneOffset(endIso, timezone), timezone_source, date_text: dateText || question || undefined, resolution };
}

export function questionTypeFromArgs(args: Record<string, unknown>): CalendarQuestionType {
  const supplied = typeof args.question_type === "string" ? args.question_type.trim() : "";
  if (["day_brief", "important_meetings", "next_meeting", "schedule_summary"].includes(supplied)) return supplied as CalendarQuestionType;
  const q = typeof args.question === "string" ? args.question.toLowerCase() : "";
  if (q.includes("next meeting")) return "next_meeting";
  if (q.includes("important") || q.includes("pay attention") || q.includes("urgent") || q.includes("matter")) return "important_meetings";
  if (q.includes("look like") || q.includes("schedule") || q.includes("calendar")) return "schedule_summary";
  return "day_brief";
}

export function scoreCalendarEvent(event: CalendarMetadataEvent, now = new Date()): { score: number; label: string; reasons: string[] } {
  const subject = typeof event.subject === "string" ? event.subject.toLowerCase() : "";
  const counts = attendeeCounts(event.attendees);
  const reasons: string[] = [];
  let score = 0;
  const matched = importantCalendarKeywords.filter((keyword) => subject.includes(keyword));
  if (matched.length) { score += 3; reasons.push("the subject contains business-priority language"); }
  if (counts.external_attendee_count > 0) { score += 2; reasons.push("it includes external attendees"); }
  if (counts.attendee_count >= 3) { score += 1; reasons.push("it has multiple attendees"); }
  const showAs = typeof event.showAs === "string" ? event.showAs.toLowerCase() : "";
  if (["busy", "oof", "workingelsewhere"].includes(showAs)) { score += 1; reasons.push("your calendar marks the time as busy"); }
  const response = event.responseStatus && typeof event.responseStatus === "object" ? String((event.responseStatus as Record<string, unknown>).response || "").toLowerCase() : "";
  if (response && response !== "declined") { score += 1; reasons.push(`your response is ${safeCalendarText(response, "set")}`); }
  const startValue = eventDateValue(event.start);
  if (startValue) {
    const startsAt = new Date(startValue).getTime();
    if (Number.isFinite(startsAt) && startsAt >= now.getTime() && startsAt - now.getTime() <= 4 * 60 * 60 * 1000) {
      score += 1;
      reasons.push("it is coming up soon");
    }
  }
  if (hasSafeLocation(event.location) || event.isOnlineMeeting === true) { reasons.push("it has a meeting location or online-meeting flag"); }
  const label = score >= 4 ? "likely needs attention" : score >= 2 ? "may be worth attention" : "routine metadata signal";
  return { score, label, reasons: reasons.slice(0, 4) };
}

export function summarizeCalendarEvent(event: CalendarMetadataEvent, timeZone = "America/New_York"): Record<string, unknown> {
  const counts = attendeeCounts(event.attendees);
  const importance = scoreCalendarEvent(event);
  const id = typeof event.id === "string" ? event.id : JSON.stringify(event).slice(0, 120);
  const startInstant = eventInstant(event.start);
  const endInstant = eventInstant(event.end);
  return {
    event_id_hash: sha256Hex(id).slice(0, 16),
    subject: safeCalendarText(event.subject),
    organizer_display: safeOrganizerDisplay(event.organizer),
    start: startInstant ? isoWithTimezoneOffset(startInstant.toISOString(), timeZone) : eventDateValue(event.start),
    end: endInstant ? isoWithTimezoneOffset(endInstant.toISOString(), timeZone) : eventDateValue(event.end),
    start_display: startInstant ? formatDisplayTime(startInstant, timeZone) : eventDisplayValue(event.start, timeZone),
    end_display: endInstant ? formatDisplayTime(endInstant, timeZone) : eventDisplayValue(event.end, timeZone),
    display_timezone: normalizeTimezoneName(timeZone),
    show_as: typeof event.showAs === "string" ? safeCalendarText(event.showAs, "unknown") : null,
    response_status: event.responseStatus && typeof event.responseStatus === "object" ? safeCalendarText((event.responseStatus as Record<string, unknown>).response, "unknown") : null,
    attendee_count: counts.attendee_count,
    external_attendee_count: counts.external_attendee_count,
    has_safe_location_label: hasSafeLocation(event.location),
    is_online_meeting: event.isOnlineMeeting === true,
    importance,
  };
}

export function appendCalendarQueryAudit(event: Record<string, unknown>): void {
  mkdirSync(CALENDAR_QUERY_AUDIT_PATH.split("/").slice(0, -1).join("/"), { recursive: true, mode: 0o700 });
  const metadata = redact({ timestamp_utc: isoNow(), source: "emclaw_m365_calendar_query", tokens_logged: false, raw_slack_ids_logged: false, raw_emails_logged: false, m365_content_accessed: false, ...event });
  appendFileSync(CALENDAR_QUERY_AUDIT_PATH, `${JSON.stringify(metadata)}\n`, { mode: 0o600 });
}

export function calendarWindowLabel(window: { window: string; resolution: string; date_text?: string }): string {
  const res = window.resolution || window.window || "";
  const weekday = res.match(/^(this|next|last)_([a-z]{3})/);
  if (weekday) {
    const days: Record<string, string> = { sun: "Sunday", mon: "Monday", tue: "Tuesday", wed: "Wednesday", thu: "Thursday", fri: "Friday", sat: "Saturday" };
    if (days[weekday[2]]) return `${weekday[1]} ${days[weekday[2]]}`;
  }
  const map: Record<string, string> = {
    today: "today", tomorrow: "tomorrow", yesterday: "yesterday",
    this_afternoon: "this afternoon", next_24_hours: "the next 24 hours",
    this_week: "this week", last_week: "last week", last_7_days: "the last 7 days",
    previous_business_week: "last business week", monday_to_friday: "Monday to Friday",
    absolute_day: window.date_text || "that day", week_of: window.date_text || "that week",
    explicit_start_end: window.date_text || "that range", relative_window: (window.window || "").replace(/_/g, " "),
  };
  if (map[res]) return map[res];
  return (window.window === "explicit_range" ? (window.date_text || res) : window.window || res).replace(/_/g, " ");
}

export function formatCalendarQueryAnswer(caller: Caller, args: Record<string, unknown>, events: CalendarMetadataEvent[], graphCalled: boolean, auditExtra: Record<string, unknown> = {}): ToolResult {
  const preferences = ensurePreferenceRecord(caller);
  const displayName = activePreferenceValue(preferences, "display_name_preference");
  const questionType = questionTypeFromArgs(args);
  const window = calendarWindowFromArgs(args);
  const summaries = events.map((event) => summarizeCalendarEvent(event, window.timezone));
  const sorted = [...summaries].sort((a, b) => ((b.importance as { score?: number }).score || 0) - ((a.importance as { score?: number }).score || 0));
  const greeting = displayName ? `${safeCalendarText(displayName, "")}, ` : "";
  const windowLabel = calendarWindowLabel(window);
  let answer = summaries.length === 0 ? `${greeting}I found no calendar items on your calendar for ${windowLabel}.` : `${greeting}I found ${summaries.length} calendar item${summaries.length === 1 ? "" : "s"} for ${windowLabel}.`;
  if (summaries.length === 0) {
    answer += " Nothing stands out from the basic calendar details I checked.";
  } else if (questionType === "next_meeting") {
    const now = new Date();
    const future = summaries.filter((event) => {
      const startMs = Date.parse(String(event.start || ""));
      return Number.isFinite(startMs) && startMs >= now.getTime();
    });
    const next = [...(future.length ? future : summaries)].sort((a, b) => String(a.start || "").localeCompare(String(b.start || "")))[0];
    answer += ` Your next listed calendar item is ${next.subject} at ${next.start_display || next.start || "the listed start time"} ${normalizeTimezoneName(window.timezone)}.`;
  } else {
    const top = sorted[0];
    const imp = top.importance as { label?: string; reasons?: string[] };
    answer += ` The item most likely to matter is ${top.subject} at ${top.start_display || top.start || "the listed start time"} ${normalizeTimezoneName(window.timezone)}; it ${imp.label || "has basic-detail signals"}.`;
    if (Array.isArray(imp.reasons) && imp.reasons.length) answer += ` Why: ${imp.reasons.join("; ")}.`;
  }
  answer += " I only checked event basics like title, time, organizer, and busy/free status. I did not read event notes, attachments, attendee emails, or meeting links.";
  appendCalendarQueryAudit({ caller_user_ref: caller.user_ref, operation: "m365_calendar_query", question_type: questionType, window: window.window, resolved_start: window.start, resolved_end: window.end, date_text: window.date_text, date_resolution: window.resolution, timezone: window.timezone, timezone_source: window.timezone_source, resolved_start_with_offset: window.start_with_offset, resolved_end_with_offset: window.end_with_offset, display_timezone: window.timezone, result_count: summaries.length, graph_called: graphCalled, graph_call_scope: graphCalled ? "calendar_metadata_only" : "none", content_access: "metadata_only", body_read: false, attachments_read: false, calendar_write: false, tokens_logged: false, raw_emails_logged: false, raw_slack_ids_logged: false, ...auditExtra });
  return okResult({ action: "m365_calendar_query", operation_id: questionType === "day_brief" || questionType === "schedule_summary" ? "calendar.day_brief" : "calendar.question_answer", user_ref: caller.user_ref, window, question_type: questionType, answer_text: answer, events: summaries, allowed_fields: calendarQuerySelectFields, excluded_fields: ["body", "bodyPreview", "uniqueBody", "attachments", "attendee email addresses", "online meeting join URL"], graph_called: graphCalled, graph_call_scope: graphCalled ? "calendar_metadata_only" : "none", content_access: "metadata_only", calendar_body_read: false, calendar_attachments_read: false, calendar_write: false, raw_attendee_emails_exposed: false, meeting_links_exposed: false, workflow_created: false, ...auditExtra });
}

export const reservedOidcScopes = new Set(["offline_access", "openid", "profile"]);

export type TokenLoadResult = { tokenPath: string; wrapper: Record<string, unknown>; tokenResponse: Record<string, unknown> };
export type GraphFetchResult = { ok: boolean; status: number; data?: unknown; errorCategory?: string; pageCount?: number; nextLinkFollowed?: boolean };
export type GraphCalendarResult = { events: CalendarMetadataEvent[]; meta: Record<string, unknown> };
export type GraphCalendarFetchAttempt = { ok: boolean; status: number; events: CalendarMetadataEvent[]; meta: Record<string, unknown> };

export function loadTokenWrapperForGraph(userRef: string): TokenLoadResult {
  const tokenPath = `${TOKEN_DIR}/${tokenFileNameFor(userRef)}`;
  const wrapper = readJsonFile(tokenPath);
  const root = wrapper && typeof wrapper === "object" ? wrapper as Record<string, unknown> : {};
  const tokenResponse = root.token_response && typeof root.token_response === "object" ? root.token_response as Record<string, unknown> : root;
  if (typeof tokenResponse.access_token !== "string" || tokenResponse.access_token.trim() === "") throw new Error("access_token_missing");
  return { tokenPath, wrapper: root, tokenResponse };
}

export function loadM365ClientConfig(): { tenant_id: string; client_id: string; scopes: string[] } {
  const data = readJsonFile(M365_CLIENT_CONFIG_PATH);
  const record = data && typeof data === "object" ? data as Record<string, unknown> : {};
  const rawScopes = Array.isArray(record.scopes) ? record.scopes : (typeof record.scopes === "string" ? record.scopes.split(/\s+/) : []);
  const scopes = rawScopes.filter((scope): scope is string => typeof scope === "string" && scope.trim() !== "" && reservedOidcScopes.has(scope.trim()) === false).map((scope) => scope.trim());
  if (typeof record.tenant_id !== "string" || record.tenant_id.trim() === "" || typeof record.client_id !== "string" || record.client_id.trim() === "" || scopes.length === 0) throw new Error("m365_client_config_incomplete");
  return { tenant_id: record.tenant_id.trim(), client_id: record.client_id.trim(), scopes };
}

export function updateTokenAfterRefresh(tokenPath: string, wrapper: Record<string, unknown>, tokenResponse: Record<string, unknown>): void {
  const now = isoNow();
  const updated = { ...wrapper, token_response: tokenResponse, updated_utc: now, last_refresh_utc: now, last_refresh_method: "oauth2_refresh_token", tokens_printed: false, graph_called: true, m365_content_accessed: false };
  writePrivateJsonAtomic(tokenPath, updated);
}

export async function refreshAccessToken(load: TokenLoadResult): Promise<boolean> {
  const refreshToken = load.tokenResponse.refresh_token;
  if (typeof refreshToken !== "string" || refreshToken.trim() === "") throw new Error("refresh_token_missing");
  const cfg = loadM365ClientConfig();
  const body = new URLSearchParams();
  body.set("client_id", cfg.client_id);
  body.set("grant_type", "refresh_token");
  body.set("refresh_token", refreshToken);
  body.set("scope", cfg.scopes.join(" "));
  const fetchImpl = (globalThis as unknown as { fetch?: (input: string, init?: Record<string, unknown>) => Promise<{ ok: boolean; status: number; json: () => Promise<unknown> }> }).fetch;
  if (typeof fetchImpl !== "function") throw new Error("fetch_unavailable");
  const response = await fetchImpl(`https://login.microsoftonline.com/${encodeURIComponent(cfg.tenant_id)}/oauth2/v2.0/token`, { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" }, body: body.toString() });
  if (response.ok !== true) throw new Error(response.status === 400 || response.status === 401 ? "refresh_token_expired_or_invalid" : `token_refresh_http_${response.status}`);
  const data = await response.json();
  if (!data || typeof data !== "object" || typeof (data as Record<string, unknown>).access_token !== "string") throw new Error("refresh_access_token_missing");
  updateTokenAfterRefresh(load.tokenPath, load.wrapper, data as Record<string, unknown>);
  Object.assign(load.tokenResponse, data as Record<string, unknown>);
  return true;
}

export function accessTokenFromLoad(load: TokenLoadResult): string {
  const accessToken = load.tokenResponse.access_token;
  if (typeof accessToken !== "string" || accessToken.trim() === "") throw new Error("access_token_missing");
  return accessToken;
}

export function graphCalendarViewUrl(window: { start: string; end: string }, calendarId?: string): string {
  const params = new URLSearchParams();
  params.set("startDateTime", "start_with_offset" in window && typeof (window as Record<string, unknown>).start_with_offset === "string" ? String((window as Record<string, unknown>).start_with_offset) : window.start);
  params.set("endDateTime", "end_with_offset" in window && typeof (window as Record<string, unknown>).end_with_offset === "string" ? String((window as Record<string, unknown>).end_with_offset) : window.end);
  params.set("$select", calendarQuerySelectFields.join(","));
  params.set("$orderby", "start/dateTime");
  params.set("$top", "50");
  const path = calendarId ? `/me/calendars/${encodeURIComponent(calendarId)}/calendarView` : "/me/calendarView";
  return `https://graph.microsoft.com/v1.0${path}?${params.toString()}`;
}

export async function graphPagedFetch(accessToken: string, initialUrl: string, window: { timezone?: string }): Promise<GraphFetchResult> {
  const fetchImpl = (globalThis as unknown as { fetch?: (input: string, init?: Record<string, unknown>) => Promise<{ ok: boolean; status: number; json: () => Promise<unknown> }> }).fetch;
  if (typeof fetchImpl !== "function") throw new Error("fetch_unavailable");
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 15000);
  try {
    const values: unknown[] = [];
    let url: string | undefined = initialUrl;
    let status = 0;
    let data: unknown = {};
    let pageCount = 0;
    while (url && pageCount < 10) {
      const response = await fetchImpl(url, { method: "GET", headers: { Authorization: `Bearer ${accessToken}`, Accept: "application/json", Prefer: `outlook.timezone="${graphWindowsTimezoneFor(window.timezone || "America/New_York")}"` }, signal: controller.signal });
      status = response.status;
      data = await response.json().catch(() => ({}));
      if (response.ok !== true) return { ok: false, status, data, pageCount: pageCount + 1, nextLinkFollowed: pageCount > 0 };
      const record = data && typeof data === "object" ? data as Record<string, unknown> : {};
      if (Array.isArray(record.value)) values.push(...record.value);
      url = typeof record["@odata.nextLink"] === "string" ? record["@odata.nextLink"] : undefined;
      pageCount += 1;
    }
    const lastRecord = data && typeof data === "object" ? data as Record<string, unknown> : {};
    return { ok: true, status, data: { ...lastRecord, value: values }, pageCount, nextLinkFollowed: pageCount > 1 };
  } finally {
    clearTimeout(timeout);
  }
}

export async function graphCalendarFetch(accessToken: string, window: { start: string; end: string; timezone?: string; start_with_offset?: string; end_with_offset?: string }, calendarId?: string): Promise<GraphFetchResult> {
  const result = await graphPagedFetch(accessToken, graphCalendarViewUrl(window, calendarId), window);
  if (result.ok !== true || !calendarId) return result;
  const data = result.data && typeof result.data === "object" ? result.data as Record<string, unknown> : {};
  const value = Array.isArray(data.value) ? data.value.map((item) => item && typeof item === "object" ? { ...(item as Record<string, unknown>), source_calendar_hash: sha256Hex(calendarId).slice(0, 16), source: "own_calendar" } : item) : [];
  return { ...result, data: { ...data, value } };
}

export function graphCalendarListUrl(): string {
  // No $select: return full calendar objects so isSharedWithMe is present.
  // (A prior `$select=id,isSharedWithMe` produced HTTP 400 due to comma encoding;
  // omitting $select avoids that while still exposing isSharedWithMe for ownership.)
  const params = new URLSearchParams();
  params.set("$top", "50");
  return `https://graph.microsoft.com/v1.0/me/calendars?${params.toString()}`;
}

export function graphCalendarGroupsListUrl(): string {
  const params = new URLSearchParams();
  params.set("$select", "id");
  params.set("$top", "50");
  return `https://graph.microsoft.com/v1.0/me/calendarGroups?${params.toString()}`;
}

export function graphCalendarsInGroupUrl(groupId: string): string {
  // No $select so isSharedWithMe is returned for calendars in non-default groups too.
  const params = new URLSearchParams();
  params.set("$top", "50");
  return `https://graph.microsoft.com/v1.0/me/calendarGroups/${encodeURIComponent(groupId)}/calendars?${params.toString()}`;
}

export async function graphCalendarListFetch(accessToken: string): Promise<GraphFetchResult> {
  return graphPagedFetch(accessToken, graphCalendarListUrl(), { timezone: "America/New_York" });
}

export function calendarEventsFromGraphData(data: unknown): CalendarMetadataEvent[] {
  const value = data && typeof data === "object" ? (data as { value?: unknown }).value : undefined;
  return Array.isArray(value) ? value.filter((item) => item && typeof item === "object") as CalendarMetadataEvent[] : [];
}

// Strict half-open overlap test [windowStart, windowEnd): keeps an event only if it actually spans
// into the window. Graph calendarView returns events that merely TOUCH the boundary (e.g. an all-day
// item ending exactly at the next day's midnight = the query's start), which previously leaked an
// adjacent-day all-day event (e.g. June 12 holiday) into a June 13 query. Unparseable times are kept.
export function eventOverlapsWindow(event: CalendarMetadataEvent, windowStartIso: string, windowEndIso: string): boolean {
  const windowStart = Date.parse(windowStartIso);
  const windowEnd = Date.parse(windowEndIso);
  if (!Number.isFinite(windowStart) || !Number.isFinite(windowEnd)) return true;
  const startInstant = eventInstant(event.start);
  const endInstant = eventInstant(event.end);
  const startMs = startInstant ? startInstant.getTime() : null;
  const endMs = endInstant ? endInstant.getTime() : null;
  if (startMs === null && endMs === null) return true; // do not silently drop events we cannot place
  const effectiveStart = startMs ?? (endMs as number);
  const effectiveEnd = endMs ?? (startMs as number);
  return effectiveStart < windowEnd && effectiveEnd > windowStart;
}

export function calendarEventsInWindow(events: CalendarMetadataEvent[], window: { start: string; end: string }): CalendarMetadataEvent[] {
  return events.filter((event) => eventOverlapsWindow(event, window.start, window.end));
}

export function ownCalendarIdsFromGraphData(data: unknown): string[] {
  return calendarDescriptorsFromGraphData(data).filter((d) => d.isSharedWithMe !== true).map((d) => d.id);
}

export function calendarDescriptorsFromGraphData(data: unknown): { id: string; isSharedWithMe: boolean }[] {
  const value = data && typeof data === "object" ? (data as { value?: unknown }).value : undefined;
  if (!Array.isArray(value)) return [];
  const seen = new Set<string>();
  const out: { id: string; isSharedWithMe: boolean }[] = [];
  for (const item of value) {
    const record = item && typeof item === "object" ? item as Record<string, unknown> : {};
    const id = typeof record.id === "string" ? record.id : "";
    if (!id || seen.has(id)) continue;
    seen.add(id);
    out.push({ id, isSharedWithMe: record.isSharedWithMe === true });
  }
  return out;
}

export function calendarGroupIdsFromGraphData(data: unknown): string[] {
  const value = data && typeof data === "object" ? (data as { value?: unknown }).value : undefined;
  if (!Array.isArray(value)) return [];
  const ids: string[] = [];
  for (const item of value) {
    const record = item && typeof item === "object" ? item as Record<string, unknown> : {};
    const id = typeof record.id === "string" ? record.id : "";
    if (id) ids.push(id);
  }
  return Array.from(new Set(ids));
}

export function teamIntentFromArgs(args: Record<string, unknown>): TeamAvailabilityIntent {
  const supplied = typeof args.intent === "string" ? args.intent.trim() : "";
  if (["team_free_busy", "find_mutual_time", "only_my_events", "shared_events", "classify_event"].includes(supplied)) return supplied as TeamAvailabilityIntent;
  const q = typeof args.question === "string" ? args.question.toLowerCase() : "";
  if (q.includes("only my") || q.includes("my events")) return "only_my_events";
  if (q.includes("shared") || q.includes("team event")) return "shared_events";
  if (q.includes("find a time") || q.includes("both available") || q.includes("mutual")) return "find_mutual_time";
  if (q.includes("my event or a team event") || q.includes("look like")) return "classify_event";
  return "team_free_busy";
}

export function safeTeamRefsFromArgs(args: Record<string, unknown>, caller: Caller): string[] {
  const roster = loadRoster().filter((entry) => entry.status === "active" && entry.role === "member");
  const allowed = new Set(roster.map((entry) => entry.user_ref));
  const raw = Array.isArray(args.team_refs) ? args.team_refs : [];
  const requested = raw.filter((item): item is string => typeof item === "string" && allowed.has(item));
  const refs = requested.length ? requested : roster.map((entry) => entry.user_ref);
  if (caller.role === "member" && !refs.includes(caller.user_ref)) refs.push(caller.user_ref);
  return Array.from(new Set(refs)).sort();
}

export function loadTeamScheduleTargets(teamRefs: string[]): TeamScheduleTarget[] {
  const data = readJsonFile(PRIVATE_USER_DIRECTORY_PATH);
  const users = data && typeof data === "object" && Array.isArray((data as { users?: unknown }).users) ? (data as { users: unknown[] }).users : [];
  const byRef = new Map<string, TeamScheduleTarget>();
  for (const item of users) {
    const user = item && typeof item === "object" ? item as Record<string, unknown> : {};
    const userRef = typeof user.user_ref === "string" ? user.user_ref : "";
    const accounts = user.app_expected_accounts && typeof user.app_expected_accounts === "object" ? user.app_expected_accounts as Record<string, unknown> : {};
    const schedule = typeof accounts.microsoft365 === "string" ? accounts.microsoft365.trim() : "";
    if (userRef && schedule) byRef.set(userRef, { user_ref: userRef, schedule, schedule_hash: sha256Hex(schedule.toLowerCase()).slice(0, 16) });
  }
  return teamRefs.map((userRef) => byRef.get(userRef)).filter((target): target is TeamScheduleTarget => Boolean(target));
}

export async function graphTeamScheduleFetch(accessToken: string, window: { start: string; end: string; timezone?: string; start_with_offset?: string; end_with_offset?: string }, schedules: string[]): Promise<GraphFetchResult> {
  const fetchImpl = (globalThis as unknown as { fetch?: (input: string, init?: Record<string, unknown>) => Promise<{ ok: boolean; status: number; json: () => Promise<unknown> }> }).fetch;
  if (typeof fetchImpl !== "function") throw new Error("fetch_unavailable");
  const timezone = graphWindowsTimezoneFor(window.timezone || "America/New_York");
  const payload = {
    schedules,
    startTime: { dateTime: window.start_with_offset || window.start, timeZone: timezone },
    endTime: { dateTime: window.end_with_offset || window.end, timeZone: timezone },
    availabilityViewInterval: 60,
  };
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 15000);
  try {
    const response = await fetchImpl("https://graph.microsoft.com/v1.0/me/calendar/getSchedule", { method: "POST", headers: { Authorization: `Bearer ${accessToken}`, Accept: "application/json", "Content-Type": "application/json", Prefer: `outlook.timezone="${timezone}"` }, body: JSON.stringify(payload), signal: controller.signal });
    const data = await response.json().catch(() => ({}));
    return { ok: response.ok === true, status: response.status, data };
  } finally {
    clearTimeout(timeout);
  }
}

export function scheduleEventsFromGetScheduleData(data: unknown, targets: TeamScheduleTarget[], timeZone: string): CalendarMetadataEvent[] {
  const value = data && typeof data === "object" ? (data as { value?: unknown }).value : undefined;
  if (!Array.isArray(value)) return [];
  const events: CalendarMetadataEvent[] = [];
  value.forEach((entry, index) => {
    const target = targets[index];
    if (!target || !entry || typeof entry !== "object") return;
    const items = Array.isArray((entry as Record<string, unknown>).scheduleItems) ? (entry as { scheduleItems: unknown[] }).scheduleItems : [];
    for (const item of items) {
      const record = item && typeof item === "object" ? item as Record<string, unknown> : {};
      const status = typeof record.status === "string" ? record.status : "busy";
      if (["free", "workingElsewhere", "unknown"].includes(status)) continue;
      const start = record.start && typeof record.start === "object" ? record.start as Record<string, unknown> : {};
      const end = record.end && typeof record.end === "object" ? record.end as Record<string, unknown> : {};
      const startValue = typeof start.dateTime === "string" ? start.dateTime : null;
      const endValue = typeof end.dateTime === "string" ? end.dateTime : null;
      if (!startValue || !endValue) continue;
      events.push({
        id: `freebusy-${sha256Hex(`${target.user_ref}:${startValue}:${endValue}:${status}`).slice(0, 16)}`,
        owner_user_ref: target.user_ref,
        source: "teammate_freebusy",
        showAs: status,
        start: { dateTime: startValue, timeZone: typeof start.timeZone === "string" ? start.timeZone : timeZone },
        end: { dateTime: endValue, timeZone: typeof end.timeZone === "string" ? end.timeZone : timeZone },
        subject: null,
        attendees: [],
      });
    }
  });
  return events;
}

export function teamScheduleEventsInWindow(data: unknown, targets: TeamScheduleTarget[], window: { start: string; end: string; timezone?: string }): CalendarMetadataEvent[] {
  return calendarEventsInWindow(scheduleEventsFromGetScheduleData(data, targets, window.timezone || "America/New_York"), window);
}

export function classifyCalendarOwnership(event: CalendarMetadataEvent, caller: Caller): Record<string, unknown> {
  const source = typeof event.source === "string" ? event.source : "own_calendar";
  const groupSource = source === "group_calendar";
  const teammateFreebusy = source === "teammate_freebusy";
  const connectedBasic = source === "connected_member_calendar_basic";
  const ownerRef = typeof event.owner_user_ref === "string" ? event.owner_user_ref : caller.user_ref;
  const counts = attendeeCounts(event.attendees);
  const internalCount = internalAttendeeCount(event.attendees);
  const selfIncluded = ownerRef === caller.user_ref || source === "own_calendar";
  const hasTeamSignals = groupSource || internalCount > 1 || Boolean(event.team_shared === true);
  let owner_scope: "personal_self" | "team_shared" | "group_calendar" | "other_member_only" | "unknown" = "unknown";
  if (groupSource) owner_scope = "group_calendar";
  else if (teammateFreebusy || (connectedBasic && ownerRef !== caller.user_ref)) owner_scope = "other_member_only";
  else if (selfIncluded && hasTeamSignals) owner_scope = "team_shared";
  else if (selfIncluded) owner_scope = "personal_self";
  let participation_scope: "self_only" | "self_and_team" | "team_not_self" | "external_mixed" | "unknown" = "unknown";
  if (counts.external_attendee_count > 0 && (internalCount > 0 || selfIncluded)) participation_scope = "external_mixed";
  else if (owner_scope === "other_member_only" || teammateFreebusy) participation_scope = "team_not_self";
  else if (owner_scope === "team_shared" || owner_scope === "group_calendar") participation_scope = selfIncluded ? "self_and_team" : "team_not_self";
  else if (owner_scope === "personal_self") participation_scope = "self_only";
  const organizerHashInput = JSON.stringify(redact(event.organizer ?? null));
  return {
    owner_scope,
    participation_scope,
    source: groupSource ? "group_calendar" : teammateFreebusy ? "teammate_freebusy" : connectedBasic ? "connected_member_calendar_basic" : "own_calendar",
    event_id_hash: typeof event.id === "string" ? sha256Hex(event.id).slice(0, 16) : null,
    subject: owner_scope === "other_member_only" ? null : safeCalendarText(event.subject),
    subject_hash: typeof event.subject === "string" ? sha256Hex(event.subject).slice(0, 16) : null,
    organizer_display: owner_scope === "other_member_only" ? null : safeOrganizerDisplay(event.organizer),
    organizer_hash: organizerHashInput && organizerHashInput !== "null" ? sha256Hex(organizerHashInput).slice(0, 16) : null,
    start: eventDateValue(event.start),
    end: eventDateValue(event.end),
    show_as: typeof event.showAs === "string" ? safeCalendarText(event.showAs, "unknown") : null,
    internal_attendee_count: internalCount,
    external_attendee_flag: counts.external_attendee_count > 0,
    uncertain: owner_scope === "unknown" || participation_scope === "unknown",
  };
}

export type SafeUserDisplayNameOptions = { selfLabel?: string; fallbackLabel?: string };

export function safePrivateDirectoryDisplayName(userRef: string): string | null {
  const rosterAllowsName = loadRoster().some((entry) => entry.user_ref === userRef && entry.status === "active");
  if (!rosterAllowsName) return null;
  try {
    const data = readJsonFile(PRIVATE_USER_DIRECTORY_PATH);
    const users = data && typeof data === "object" && Array.isArray((data as { users?: unknown }).users) ? (data as { users: unknown[] }).users : [];
    for (const item of users) {
      const user = item && typeof item === "object" ? item as Record<string, unknown> : {};
      if (user.user_ref !== userRef) continue;
      const displayName = typeof user.display_name === "string" ? user.display_name.trim() : "";
      return displayName || null;
    }
  } catch { /* directory unavailable; fall back below */ }
  return null;
}

// Resolve a friendly, non-identifying display name for user-facing basic-status answers.
// Prefer the private directory display_name for active roster users, then the user's
// self-chosen display_name_preference, then operator-approved fallback labels.
// New/unapproved teammates stay generic until onboarding or preference state approves
// a name. Never returns raw Slack IDs, user_refs, emails, or hash labels.
export function safeUserDisplayName(userRef: string, caller: Caller, options: SafeUserDisplayNameOptions = {}): string {
  const selfLabel = options.selfLabel || "you";
  const fallbackLabel = options.fallbackLabel || "a teammate";
  if (userRef === caller.user_ref) return selfLabel;
  const privateDirectoryName = safePrivateDirectoryDisplayName(userRef);
  if (privateDirectoryName) return safeCalendarText(privateDirectoryName, fallbackLabel);
  try {
    const data = readJsonFile(preferencePathFor(userRef));
    if (data && typeof data === "object") {
      const name = activePreferenceValue(data as Record<string, unknown>, "display_name_preference");
      if (name) return safeCalendarText(name, fallbackLabel);
    }
  } catch { /* no preference file yet */ }
  const approvedName = APPROVED_SAFE_DISPLAY_NAMES_BY_REF[userRef];
  if (approvedName) return safeCalendarText(approvedName, fallbackLabel);
  return fallbackLabel;
}

export function teammateDisplayName(userRef: string, caller: Caller): string {
  return safeUserDisplayName(userRef, caller, { fallbackLabel: "a teammate" });
}

export function freeBusyRowsForTeam(teamRefs: string[], events: CalendarMetadataEvent[], caller: Caller): Record<string, unknown>[] {
  return teamRefs.map((userRef) => {
    const userEvents = events.filter((event) => (typeof event.owner_user_ref === "string" ? event.owner_user_ref : caller.user_ref) === userRef);
    const busy = userEvents.some((event) => {
      const showAs = typeof event.showAs === "string" ? event.showAs.toLowerCase() : "busy";
      return showAs !== "free" && showAs !== "tentative";
    });
    return { user_ref: userRef, display_name: teammateDisplayName(userRef, caller), availability: busy ? "busy" : "free", source: userRef === caller.user_ref ? "own_calendar" : "teammate_freebusy", event_details_exposed: false };
  });
}

export function appendTeamCalendarAudit(event: Record<string, unknown>): void {
  mkdirSync(TEAM_CALENDAR_AUDIT_PATH.split("/").slice(0, -1).join("/"), { recursive: true, mode: 0o700 });
  const metadata = redact({ timestamp_utc: isoNow(), source: "emclaw_m365_team_availability", tokens_logged: false, raw_slack_ids_logged: false, raw_emails_logged: false, m365_content_accessed: false, calendar_write: false, ...event });
  appendFileSync(TEAM_CALENDAR_AUDIT_PATH, `${JSON.stringify(metadata)}\n`, { mode: 0o600 });
}

export function formatTeamAvailabilityAnswer(caller: Caller, args: Record<string, unknown>, events: CalendarMetadataEvent[], graphCalled: boolean, auditExtra: Record<string, unknown> = {}): ToolResult {
  const intent = teamIntentFromArgs(args);
  const window = calendarWindowFromArgs(args);
  const teamRefs = safeTeamRefsFromArgs(args, caller);
  const classified = events.map((event) => classifyCalendarOwnership(event, caller));
  const freeBusy = freeBusyRowsForTeam(teamRefs, events, caller);
  let filtered = classified;
  if (intent === "only_my_events") filtered = classified.filter((event) => event.owner_scope === "personal_self");
  if (intent === "shared_events") filtered = classified.filter((event) => event.owner_scope === "team_shared" || event.owner_scope === "group_calendar");
  const freeCount = freeBusy.filter((row) => row.availability === "free").length;
  let answer = "";
  if (intent === "only_my_events") {
    answer = filtered.length === 0 ? "I found no events that look personal-only for you in that window." : `I found ${filtered.length} event${filtered.length === 1 ? "" : "s"} that look personal-only for you. I left out other members' private events.`;
  } else if (intent === "shared_events" || intent === "classify_event") {
    answer = filtered.length === 0 ? "I did not find events that safely look shared with the team in that window." : `I found ${filtered.length} event${filtered.length === 1 ? "" : "s"} that safely look shared with the team.`;
  } else {
    const statuses = freeBusy.map((row) => {
      const name = typeof row.display_name === "string" && row.display_name ? row.display_name : "a teammate";
      const verb = row.user_ref === caller.user_ref ? "are" : "is";
      return `${name} ${verb} ${row.availability === "free" ? "free" : "busy"}`;
    });
    const list = statuses.length <= 1 ? (statuses[0] || "no one is listed") : `${statuses.slice(0, -1).join(", ")} and ${statuses[statuses.length - 1]}`;
    const summaryAnswer = `${list.charAt(0).toUpperCase()}${list.slice(1)} in that window, from the free/busy details I checked (${freeCount} of ${teamRefs.length} free).`;
    answer = summaryAnswer;
  }
  answer += " I only checked basic calendar details like title, time, organizer, busy/free status, and whether it appears shared. I did not read event notes, attachments, attendee emails, or meeting links.";
  appendTeamCalendarAudit({ caller_user_ref: caller.user_ref, operation: "m365_team_availability", intent, window: window.window, team_ref_count: teamRefs.length, result_count: filtered.length, graph_called: graphCalled, graph_call_scope: graphCalled ? "calendar_free_busy_basic" : "none", content_access: "basic_calendar_details_only", body_read: false, attachments_read: false, calendar_write: false, raw_attendee_emails_logged: false, meeting_links_logged: false, group_calendar_read_used: Boolean(auditExtra.group_calendar_read_used), ...auditExtra });
  return okResult({ action: "m365_team_availability", operation_id: "calendar.team_availability_basic", user_ref: caller.user_ref, intent, window, answer_text: answer, team_refs: teamRefs, free_busy: freeBusy, events: filtered, owner_scope_values: ["personal_self", "team_shared", "group_calendar", "other_member_only", "unknown"], participation_scope_values: ["self_only", "self_and_team", "team_not_self", "external_mixed", "unknown"], source_values: ["own_calendar", "teammate_freebusy", "connected_member_calendar_basic", "group_calendar", "shared_calendar", "unknown"], getSchedule_supported_or_blocked: auditExtra.getSchedule_supported_or_blocked || "blocked_schedule_addresses_not_configured", group_calendar_read_used: Boolean(auditExtra.group_calendar_read_used), group_readwrite_all_required: false, free_busy_only_default: intent === "team_free_busy" || intent === "find_mutual_time", content_access: "basic_calendar_details_only", calendar_body_read: false, calendar_attachments_read: false, calendar_write: false, raw_attendee_emails_exposed: false, meeting_links_exposed: false, workflow_created: false, graph_called: graphCalled, graph_call_scope: graphCalled ? "calendar_free_busy_basic" : "none", ...auditExtra });
}

export async function fetchTeamAvailability(caller: Caller, window: { start: string; end: string; timezone?: string; start_with_offset?: string; end_with_offset?: string }, targets: TeamScheduleTarget[]): Promise<{ events: CalendarMetadataEvent[]; meta: Record<string, unknown> }> {
  const schedules = targets.map((target) => target.schedule);
  const load = loadTokenWrapperForGraph(caller.user_ref);
  const first = await graphTeamScheduleFetch(accessTokenFromLoad(load), window, schedules);
  if (first.ok === true) {
    return {
      events: teamScheduleEventsInWindow(first.data, targets, window),
      meta: { graph_call_count: 1, final_http_status: first.status, refresh_attempted: false, refresh_succeeded: false, token_refreshed: false, getSchedule_supported_or_blocked: "getSchedule_used" },
    };
  }
  if (first.status !== 401) throw new Error(`graph_team_getSchedule_http_${first.status}`);
  await refreshAccessToken(load);
  const second = await graphTeamScheduleFetch(accessTokenFromLoad(load), window, schedules);
  if (second.ok !== true) throw new Error(`graph_team_getSchedule_http_${second.status}`);
  return {
    events: teamScheduleEventsInWindow(second.data, targets, window),
    meta: { graph_call_count: 2, initial_http_status: first.status, final_http_status: second.status, refresh_attempted: true, refresh_succeeded: true, token_refreshed: true, getSchedule_supported_or_blocked: "getSchedule_used" },
  };
}

export async function m365TeamAvailabilityForCaller(args: Record<string, unknown>, caller: Caller): Promise<ToolResult> {
  requireSelfCapable(caller);
  const syntheticEvents = Array.isArray(args.synthetic_events) ? args.synthetic_events.filter((item) => item && typeof item === "object") as CalendarMetadataEvent[] : [];
  const dryRun = args.dry_run === true || syntheticEvents.length > 0;
  // Dry-run uses synthetic events and must NOT make a live mailboxSettings call; leave args as-is.
  if (dryRun) return formatTeamAvailabilityAnswer(caller, args, syntheticEvents, false, { dry_run: true, getSchedule_supported_or_blocked: "dry_run_synthetic_free_busy_only" });
  // Resolve the caller's timezone so team availability date phrases follow the user, like the
  // calendar/email tools (precedence: explicit arg > timezone_preference > mailboxSettings > default).
  const tzResolved = await timezoneAwareArgs(args, caller);
  args = tzResolved.args;
  const tzMeta = { timezone_source: tzResolved.source, mailbox_timezone_status: tzResolved.status };
  const teamRefs = safeTeamRefsFromArgs(args, caller);
  const targets = loadTeamScheduleTargets(teamRefs);
  if (targets.length !== teamRefs.length) {
    return formatTeamAvailabilityAnswer(caller, args, [], false, { ...tzMeta, live_blocked_reason: "team_schedule_targets_missing", getSchedule_supported_or_blocked: "blocked_schedule_addresses_not_configured", schedule_target_count: targets.length, missing_schedule_target_count: teamRefs.length - targets.length, requires_raw_email_addresses: false });
  }
  const window = calendarWindowFromArgs(args);
  try {
    const result = await fetchTeamAvailability(caller, window, targets);
    return formatTeamAvailabilityAnswer(caller, args, result.events, true, { ...result.meta, ...tzMeta, user_timezone: window.timezone, schedule_target_count: targets.length, schedule_target_hashes: targets.map((target) => target.schedule_hash), final_http_status: result.meta.final_http_status || result.meta.initial_http_status, requires_raw_email_addresses: false });
  } catch (err) {
    const errorCode = err instanceof Error ? err.message : String(err);
    return formatTeamAvailabilityAnswer(caller, args, [], true, { ...tzMeta, live_blocked_reason: errorCode, getSchedule_supported_or_blocked: errorCode, schedule_target_count: targets.length, schedule_target_hashes: targets.map((target) => target.schedule_hash), requires_raw_email_addresses: false });
  }
}

export async function fetchOwnCalendarMetadataWithAccessToken(accessToken: string, window: { start: string; end: string; timezone?: string; start_with_offset?: string; end_with_offset?: string }): Promise<GraphCalendarFetchAttempt> {
  const calendars = await graphCalendarListFetch(accessToken);
  if (calendars.ok !== true) return { ok: false, status: calendars.status, events: [], meta: { graph_calendar_list_status: calendars.status, graph_calendar_list_page_count: calendars.pageCount || 1 } };
  let pageCount = calendars.pageCount || 1;
  let nextLinkFollowed = calendars.nextLinkFollowed === true;
  let status = calendars.status;

  // Discover calendars across ALL calendar groups, not just the default group that
  // /me/calendars returns. Personal/secondary calendars in a non-default group would
  // otherwise be invisible. Shared-to-me calendars are counted but never queried.
  const descriptorById = new Map<string, { id: string; isSharedWithMe: boolean }>();
  for (const d of calendarDescriptorsFromGraphData(calendars.data)) descriptorById.set(d.id, d);
  let calendarGroupCount = 0;
  const groups = await graphPagedFetch(accessToken, graphCalendarGroupsListUrl(), { timezone: "America/New_York" });
  if (groups.ok === true) {
    pageCount += groups.pageCount || 1;
    const groupIds = calendarGroupIdsFromGraphData(groups.data);
    calendarGroupCount = groupIds.length;
    for (const groupId of groupIds) {
      const groupCals = await graphPagedFetch(accessToken, graphCalendarsInGroupUrl(groupId), { timezone: "America/New_York" });
      if (groupCals.ok !== true) continue;
      pageCount += groupCals.pageCount || 1;
      for (const d of calendarDescriptorsFromGraphData(groupCals.data)) if (!descriptorById.has(d.id)) descriptorById.set(d.id, d);
    }
  }
  const allDescriptors = Array.from(descriptorById.values());
  const ownDescriptors = allDescriptors.filter((d) => d.isSharedWithMe !== true);
  const sharedSkipped = allDescriptors.filter((d) => d.isSharedWithMe === true);
  const calendarIds = ownDescriptors.map((d) => d.id);
  const events: CalendarMetadataEvent[] = [];
  const perCalendarCounts: { calendar_hash: string; is_shared_with_me: boolean; event_count: number }[] = [];
  // Shared-to-me calendars are recorded (count -1 = not queried) so we can SEE them in
  // the audit without reading another user's private calendar content.
  for (const d of sharedSkipped) perCalendarCounts.push({ calendar_hash: sha256Hex(d.id).slice(0, 16), is_shared_with_me: true, event_count: -1 });
  const discoveryMeta = { graph_calendar_group_count: calendarGroupCount, graph_total_calendars_discovered: allDescriptors.length, graph_shared_calendars_skipped: sharedSkipped.length };

  if (calendarIds.length === 0) {
    const fallback = await graphCalendarFetch(accessToken, window);
    if (fallback.ok !== true) return { ok: false, status: fallback.status, events: [], meta: { graph_calendar_list_status: calendars.status, graph_calendar_list_page_count: pageCount, fallback_default_calendar_status: fallback.status, ...discoveryMeta } };
    const fallbackEvents = calendarEventsInWindow(calendarEventsFromGraphData(fallback.data), window);
    return {
      ok: true,
      status: fallback.status,
      events: fallbackEvents,
      meta: { graph_calendar_list_status: calendars.status, graph_calendar_count: 0, graph_calendars_queried: 1, calendar_scope: "default_calendar_fallback", ...discoveryMeta, graph_per_calendar_event_counts: perCalendarCounts, graph_page_count: pageCount + (fallback.pageCount || 1), graph_next_link_followed: nextLinkFollowed || fallback.nextLinkFollowed === true },
    };
  }
  for (const calendarId of calendarIds) {
    const result = await graphCalendarFetch(accessToken, window, calendarId);
    status = result.status;
    pageCount += result.pageCount || 1;
    nextLinkFollowed = nextLinkFollowed || result.nextLinkFollowed === true;
    if (result.ok !== true) return { ok: false, status: result.status, events, meta: { graph_calendar_list_status: calendars.status, graph_calendar_count: calendarIds.length, graph_calendars_queried: calendarIds.indexOf(calendarId), calendar_scope: "own_user_calendars", ...discoveryMeta, graph_per_calendar_event_counts: perCalendarCounts, graph_page_count: pageCount, graph_next_link_followed: nextLinkFollowed } };
    const calendarEvents = calendarEventsInWindow(calendarEventsFromGraphData(result.data), window);
    perCalendarCounts.push({ calendar_hash: sha256Hex(calendarId).slice(0, 16), is_shared_with_me: false, event_count: calendarEvents.length });
    events.push(...calendarEvents);
  }
  return {
    ok: true,
    status,
    events,
    // Per-calendar breakdown is verbose; only log it when nothing was found (the case
    // where it helps diagnose a miss). Aggregate group counts stay for ongoing observability.
    meta: { graph_calendar_list_status: calendars.status, graph_calendar_count: calendarIds.length, graph_calendars_queried: calendarIds.length, calendar_scope: "own_user_calendars", ...discoveryMeta, ...(events.length === 0 ? { graph_per_calendar_event_counts: perCalendarCounts } : {}), graph_page_count: pageCount, graph_next_link_followed: nextLinkFollowed },
  };
}

export function graphMailboxTimeZoneUrl(): string {
  return "https://graph.microsoft.com/v1.0/me/mailboxSettings?$select=timeZone";
}

export async function graphMailboxTimeZoneFetch(accessToken: string): Promise<GraphFetchResult> {
  const fetchImpl = (globalThis as unknown as { fetch?: (input: string, init?: Record<string, unknown>) => Promise<{ ok: boolean; status: number; json: () => Promise<unknown> }> }).fetch;
  if (typeof fetchImpl !== "function") throw new Error("fetch_unavailable");
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 15000);
  try {
    const response = await fetchImpl(graphMailboxTimeZoneUrl(), { method: "GET", headers: { Authorization: `Bearer ${accessToken}`, Accept: "application/json" }, signal: controller.signal });
    const data = await response.json().catch(() => ({}));
    return { ok: response.ok === true, status: response.status, data };
  } finally {
    clearTimeout(timeout);
  }
}

// Resolve the caller's own timezone from Outlook (mailboxSettings.timeZone) so date windows follow
// the user's time, not the server's. No content is read — only the timezone string. Falls back gracefully.
export async function resolveUserTimeZone(userRef: string): Promise<{ timezone: string | null; source: string; status: number | null }> {
  try {
    const load = loadTokenWrapperForGraph(userRef);
    let res = await graphMailboxTimeZoneFetch(accessTokenFromLoad(load));
    let status = res.status;
    // Retry after a token refresh on 401 (expired) AND 403 (a freshly admin-consented scope like
    // MailboxSettings.Read may only appear in a newly minted access token from the refresh token).
    if (res.ok !== true && (res.status === 401 || res.status === 403)) {
      await refreshAccessToken(load);
      res = await graphMailboxTimeZoneFetch(accessTokenFromLoad(load));
      status = res.status;
    }
    if (res.ok !== true) return { timezone: null, source: "mailbox_settings_unavailable", status };
    const raw = res.data && typeof res.data === "object" ? (res.data as Record<string, unknown>).timeZone : undefined;
    const iana = ianaFromGraphTimeZone(raw);
    if (!iana) return { timezone: null, source: "mailbox_timezone_unmapped", status };
    return { timezone: iana, source: "mailbox_settings", status };
  } catch {
    return { timezone: null, source: "mailbox_settings_error", status: null };
  }
}

// Resolve the caller's timezone with precedence: explicit request arg > stored user preference >
// Outlook mailboxSettings > server default. The preference path needs no OAuth scope.
export async function timezoneAwareArgs(args: Record<string, unknown>, caller: Caller): Promise<{ args: Record<string, unknown>; source: string; status: number | null }> {
  if (typeof args.timezone === "string" && args.timezone.trim()) return { args, source: "explicit_request", status: null };
  const preferences = ensurePreferenceRecord(caller);
  const preferredRaw = activePreferenceValue(preferences, "timezone_preference");
  const preferred = preferredRaw ? userTimeZoneToIana(preferredRaw) : null;
  if (preferred) return { args: { ...args, timezone: preferred }, source: "user_preference", status: null };
  const resolved = await resolveUserTimeZone(caller.user_ref);
  if (!resolved.timezone) return { args, source: resolved.source, status: resolved.status };
  return { args: { ...args, timezone: resolved.timezone }, source: resolved.source, status: resolved.status };
}

export async function fetchOwnCalendarMetadata(userRef: string, window: { start: string; end: string; timezone?: string; start_with_offset?: string; end_with_offset?: string }): Promise<GraphCalendarResult> {
  const load = loadTokenWrapperForGraph(userRef);
  const first = await fetchOwnCalendarMetadataWithAccessToken(accessTokenFromLoad(load), window);
  if (first.ok === true) return { events: first.events, meta: { graph_call_count: 1, initial_http_status: first.status, refresh_attempted: false, refresh_succeeded: false, token_refreshed: false, prefer_outlook_timezone_used: true, ...first.meta } };
  if (first.status !== 401) throw new Error(`graph_calendar_metadata_http_${first.status}`);
  await refreshAccessToken(load);
  const second = await fetchOwnCalendarMetadataWithAccessToken(accessTokenFromLoad(load), window);
  if (second.ok !== true) throw new Error(`graph_calendar_metadata_http_${second.status}`);
  return { events: second.events, meta: { graph_call_count: 2, initial_http_status: first.status, final_http_status: second.status, refresh_attempted: true, refresh_succeeded: true, token_refreshed: true, prefer_outlook_timezone_used: true, ...second.meta } };
}

export async function m365CalendarQueryForCaller(args: Record<string, unknown>, caller: Caller): Promise<ToolResult> {
  if (tokenFileExistsByStat(caller.user_ref) !== true) return errorResult("microsoft365_connection_not_ready_by_token_stat", "Microsoft 365 is not connected for this member by private token-file stat.");
  const tzResolved = await timezoneAwareArgs(args, caller);
  args = tzResolved.args;
  const tzMeta = { timezone_source: tzResolved.source, mailbox_timezone_status: tzResolved.status };
  const window = calendarWindowFromArgs(args);
  try {
    const result = await fetchOwnCalendarMetadata(caller.user_ref, window);
    return formatCalendarQueryAnswer(caller, args, result.events, true, { ...result.meta, ...tzMeta, user_timezone: window.timezone });
  } catch (err) {
    const errorCode = err instanceof Error ? err.message : String(err);
    const reconnectNeeded = errorCode === "refresh_token_missing" || errorCode === "refresh_token_expired_or_invalid";
    appendCalendarQueryAudit({ caller_user_ref: caller.user_ref, operation: "m365_calendar_query", question_type: questionTypeFromArgs(args), window: window.window, result_count: 0, graph_called: true, graph_call_scope: "calendar_metadata_only", content_access: "metadata_only", body_read: false, attachments_read: false, calendar_write: false, tokens_logged: false, raw_emails_logged: false, raw_slack_ids_logged: false, refresh_attempted: errorCode !== "access_token_missing", refresh_succeeded: false, reconnect_required: reconnectNeeded, error_code: errorCode });
    return errorResult(reconnectNeeded ? "microsoft365_reconnect_required" : "m365_calendar_query_failed", reconnectNeeded ? "Microsoft 365 needs to be reconnected before I can check your basic calendar details." : errorCode, { action: "m365_calendar_query", user_ref: caller.user_ref, window, question_type: questionTypeFromArgs(args), allowed_fields: calendarQuerySelectFields, excluded_fields: ["body", "bodyPreview", "uniqueBody", "attachments", "attendee email addresses", "online meeting join URL"], graph_called: true, graph_call_scope: "calendar_metadata_only", content_access: "metadata_only", calendar_body_read: false, calendar_attachments_read: false, calendar_write: false, workflow_created: false, refresh_succeeded: false, reconnect_required: reconnectNeeded, token_contents_logged_or_exposed: false });
  }
}

export function dryM365CalendarQueryForCaller(caller: Caller, args: Record<string, unknown>, events: CalendarMetadataEvent[]): ToolResult {
  return formatCalendarQueryAnswer(caller, args, events, false);
}
