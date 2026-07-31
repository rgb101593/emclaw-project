import { afterEach, describe, expect, it, vi } from "vitest";
import entry, { __test } from "./index.js";
import { getToolPluginMetadata } from "openclaw/plugin-sdk/tool-plugin";

const expectedTools = [
  "m365_status",
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
  "admin_workflow_create_for_member",
  "admin_workflow_update_for_member",
  "admin_workflow_pause_for_member",
  "admin_workflow_resume_for_member",
  "admin_workflow_terminate_for_member",
  "admin_workflow_inspect_for_member",
  "admin_email_alert_quota_get",
  "admin_email_alert_quota_set",
  "create_workflow",
  "list_all_automations",
  "admin_workflow_inventory",
  "pause_automation",
  "resume_automation",
  "terminate_automation",
  "inspect_automation",
  "approve_pending_action",
  "m365_health_all_connected_accounts",
];

describe("emclaw-m365-connector", () => {
  it("declares workflow wrapper tool metadata", () => {
    expect(getToolPluginMetadata(entry)?.tools.map((tool) => tool.name)).toEqual(expectedTools);
  });

  it("marks one-time M365 email and calendar query tools as fresh per request", () => {
    const tools = getToolPluginMetadata(entry)?.tools ?? [];
    expect(tools.find((tool) => tool.name === "m365_calendar_query")?.description).toContain("fresh tool call every time");
    expect(tools.find((tool) => tool.name === "m365_email_query")?.description).toContain("fresh tool call every time");
  });
});

describe("email alert subject dedup key", () => {
  it("matches same subject across operation phrasings and ignores casing/spacing", () => {
    const a = __test.emailAlertDedupKey({ operation: "email.new_message_alert", service: "email", criteria: { subjectContains: "FINAL_AVERY_ALERT_TEST_R2" } });
    const b = __test.emailAlertDedupKey({ operation: "email.subject_contains_alert", service: "email", criteria: { subjectContains: "final avery alert test r2" } });
    expect(a).toBe(b);
  });
  it("matches same subject whether stored in filters or criteria", () => {
    const fromFilters = __test.emailAlertDedupKey({ operation: "mail.alert_metadata_delta", filters: { subjectContains: "FINAL_AVERY_ALERT_TEST_R4" } });
    const fromCriteria = __test.emailAlertDedupKey({ operation: "mail.alert_metadata_delta", criteria: { subjectContains: "final avery alert test r4" } });
    expect(fromFilters).toBe(fromCriteria);
  });
  it("distinguishes different subjects and different alert operations", () => {
    const subjX = __test.emailAlertDedupKey({ operation: "mail.alert_metadata_delta", criteria: { subjectContains: "X" } });
    const subjY = __test.emailAlertDedupKey({ operation: "mail.alert_metadata_delta", criteria: { subjectContains: "Y" } });
    const important = __test.emailAlertDedupKey({ operation: "mail.important_email_alert_metadata", criteria: {} });
    expect(subjX).not.toBe(subjY);
    expect(subjX).not.toBe(important);
  });
});


// These tests resolve identity through the committed trusted-agent mapping
// (roster.json + PER_MEMBER_AGENT_REFS). The delivery-context path that hashes
// raw Slack IDs is exercised in production from the gateway's live bindings;
// the raw IDs are intentionally absent from this public mirror, so we drive the
// same resolver here through the hermetic agent route.
describe("trusted caller identity bridge", () => {
  it("resolves members from trusted agent context", () => {
    expect(__test.resolveRosterUserFromContext({ agentId: "emclaw-member-8d4e38d9a1cf7e80-v2" }).user_ref).toBe("slack-user-ref-8d4e38d9a1cf7e80");
    expect(__test.resolveRosterUserFromContext({ agentId: "emclaw-member-db6b1d5cccad3cbf-v2" }).user_ref).toBe("slack-user-ref-db6b1d5cccad3cbf");
  });

  it("blocks unknown and spoofed member caller identity", () => {
    expect(() => __test.resolveRosterUserFromContext({ deliveryContext: { to: "unknown-route" } })).toThrow("caller_identity_unresolved");
    expect(() => __test.resolveTrustedMemberCaller({ caller_user_ref: "admin" }, { agentId: "emclaw-member-8d4e38d9a1cf7e80-v2" })).toThrow("caller_user_ref_conflicts_with_trusted_identity");
  });

  it("preserves trusted admin route fallback", () => {
    expect(__test.resolveRosterUserFromContext({ agentId: "emclaw-admin-v2" }).user_ref).toBe("admin");
  });
});

describe("calendar date window resolution", () => {
  afterEach(() => {
    vi.useRealTimers();
  });

  it("defaults next meeting questions to an upcoming 24-hour window", () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-06-05T16:10:00.000Z"));
    const window = __test.calendarWindowFromArgs({ question: "What's my next meeting?" });
    expect(window.window).toBe("next_24_hours");
    expect(window.resolution).toBe("next_24_hours");
    expect(window.start).toBe("2026-06-05T16:10:00.000Z");
    expect(window.end).toBe("2026-06-06T16:10:00.000Z");
  });

  it("resolves this Saturday as that calendar day instead of the current week so far", () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-06-05T16:12:17.000Z"));
    const window = __test.calendarWindowFromArgs({ dateText: "this Saturday" });
    expect(window.window).toBe("explicit_range");
    expect(window.resolution).toBe("this_saturday");
    expect(window.start_with_offset).toBe("2026-06-06T00:00:00-04:00");
    expect(window.end_with_offset).toBe("2026-06-07T00:00:00-04:00");
  });

  it("resolves last Tuesday to the most recent past Tuesday, not the upcoming one", () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-06-05T16:12:17.000Z")); // Friday 2026-06-05
    const window = __test.calendarWindowFromArgs({ question: "what meeting last Tuesday" });
    expect(window.window).toBe("explicit_range");
    expect(window.resolution).toBe("last_tuesday");
    expect(window.start_with_offset).toBe("2026-06-02T00:00:00-04:00");
    expect(window.end_with_offset).toBe("2026-06-03T00:00:00-04:00");
  });

  it("resolves last Friday to seven days back when today is that weekday", () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-06-05T16:12:17.000Z")); // Friday 2026-06-05
    const window = __test.calendarWindowFromArgs({ dateText: "last Friday" });
    expect(window.resolution).toBe("last_friday");
    expect(window.start_with_offset).toBe("2026-05-29T00:00:00-04:00");
    expect(window.end_with_offset).toBe("2026-05-30T00:00:00-04:00");
  });

  it("builds per-calendar calendarView URLs and filters shared-in calendars", () => {
    const window = { start: "2026-06-06T04:00:00.000Z", end: "2026-06-07T04:00:00.000Z" };
    expect(__test.graphCalendarViewUrl(window)).toContain("/me/calendarView?");
    expect(__test.graphCalendarViewUrl(window, "calendar id")).toContain("/me/calendars/calendar%20id/calendarView?");
    // List URL omits $select so isSharedWithMe is actually returned (a prior id,isSharedWithMe $select caused HTTP 400).
    expect(new URL(__test.graphCalendarListUrl()).searchParams.get("$select")).toBeNull();
    expect(__test.graphCalendarListUrl()).not.toContain("isSharedWithMe");
    expect(__test.ownCalendarIdsFromGraphData({ value: [{ id: "own-1" }, { id: "shared-1", isSharedWithMe: true }, { id: "own-2", isSharedWithMe: false }] })).toEqual(["own-1", "own-2"]);
  });

  it("labels date windows for users with readable phrases, never 'explicit range'", () => {
    expect(__test.calendarWindowLabel({ window: "explicit_range", resolution: "last_tuesday" })).toBe("last Tuesday");
    expect(__test.calendarWindowLabel({ window: "explicit_range", resolution: "this_saturday" })).toBe("this Saturday");
    expect(__test.calendarWindowLabel({ window: "explicit_range", resolution: "next_mon" })).toBe("next Monday");
    expect(__test.calendarWindowLabel({ window: "explicit_range", resolution: "absolute_day", date_text: "May 27" })).toBe("May 27");
    expect(__test.calendarWindowLabel({ window: "today", resolution: "today" })).toBe("today");
    expect(__test.calendarWindowLabel({ window: "next_24_hours", resolution: "next_24_hours" })).toBe("the next 24 hours");
    expect(__test.calendarWindowLabel({ window: "explicit_range", resolution: "this_saturday" })).not.toContain("explicit");
  });

  it("resolves date windows in the user's timezone, not the server's (the +3 Sunday case)", () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-06-05T12:00:00.000Z")); // Friday
    // User in UTC+3: "this Sunday" must be Jun 7 00:00 +03:00, so a meeting created at Jun 7 00:00 +03 is included.
    const sun = __test.calendarWindowFromArgs({ question: "what's on sunday", timezone: "Asia/Baghdad" });
    expect(sun.resolution).toBe("this_sunday");
    expect(sun.start_with_offset).toBe("2026-06-07T00:00:00+03:00");
    expect(sun.end_with_offset).toBe("2026-06-08T00:00:00+03:00");
    // Same phrase in Eastern yields a different absolute window (the old buggy behavior, now opt-in via tz).
    const sunET = __test.calendarWindowFromArgs({ question: "what's on sunday", timezone: "America/New_York" });
    expect(sunET.start_with_offset).toBe("2026-06-07T00:00:00-04:00");
    expect(sun.start).not.toBe(sunET.start); // different UTC instants
  });

  it("computes 'today' boundaries in the user's timezone", () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-06-06T01:30:00.000Z")); // 04:30 in +3 (still Jun 6), 21:30 Jun 5 in ET
    const tzPlus3 = __test.calendarWindowFromArgs({ window: "today", timezone: "Asia/Baghdad" });
    expect(tzPlus3.start_with_offset).toBe("2026-06-06T00:00:00+03:00");
    const tzET = __test.calendarWindowFromArgs({ window: "today", timezone: "America/New_York" });
    expect(tzET.start_with_offset).toBe("2026-06-05T00:00:00-04:00");
  });

  it("accepts user timezone preferences as IANA, Windows, or UTC/GMT offsets", () => {
    expect(__test.userTimeZoneToIana("Asia/Singapore")).toBe("Asia/Singapore");
    expect(__test.userTimeZoneToIana("GMT+8")).toBe("Etc/GMT-8"); // IANA sign is inverted
    expect(__test.userTimeZoneToIana("UTC+8")).toBe("Etc/GMT-8");
    expect(__test.userTimeZoneToIana("+8")).toBe("Etc/GMT-8");
    expect(__test.userTimeZoneToIana("+08:00")).toBe("Etc/GMT-8");
    expect(__test.userTimeZoneToIana("-5")).toBe("Etc/GMT+5");
    expect(__test.userTimeZoneToIana("Turkey Standard Time")).toBe("Europe/Istanbul");
    expect(__test.userTimeZoneToIana("garbage")).toBeNull();
  });

  it("computes a GMT+8 user's 'this Sunday' window correctly via a fixed offset zone", () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-06-05T12:00:00.000Z")); // Friday
    const sun = __test.calendarWindowFromArgs({ question: "what's on sunday", timezone: "Etc/GMT-8" });
    expect(sun.resolution).toBe("this_sunday");
    expect(sun.start_with_offset).toBe("2026-06-07T00:00:00+08:00");
    expect(sun.end_with_offset).toBe("2026-06-08T00:00:00+08:00");
  });

  it("resolves teammate display names without exposing raw refs", () => {
    // The caller themselves is always "you" (no preference file read).
    expect(__test.teammateDisplayName("admin", { user_ref: "admin" } as any)).toBe("you");
    expect(__test.teammateDisplayName("slack-user-ref-8d4e38d9a1cf7e80", { user_ref: "admin" } as any)).toBe("Avery");
    expect(__test.teammateDisplayName("slack-user-ref-db6b1d5cccad3cbf", { user_ref: "admin" } as any)).toBe("Cameron");
    // An unknown ref never leaks the raw ref; falls back to a generic label.
    const unknown = __test.teammateDisplayName("slack-user-ref-does-not-exist", { user_ref: "admin" } as any);
    expect(unknown).toBe("a teammate");
    expect(unknown).not.toContain("slack-user-ref-does-not-exist");
    expect(unknown).not.toMatch(/[0-9a-f]{16}/);
    expect(typeof unknown).toBe("string");
    expect(unknown.length).toBeGreaterThan(0);
  });

  it("keeps newly onboarded teammate labels generic and idempotent until approved", () => {
    const caller = { user_ref: "admin" } as any;
    const newTeammateRef = "slack-user-ref-newteammate001122334455";
    const labels = Array.from({ length: 3 }, () => __test.teammateDisplayName(newTeammateRef, caller));
    expect(labels).toEqual(["a teammate", "a teammate", "a teammate"]);
    expect(labels.join(" ")).not.toContain(newTeammateRef);
    expect(labels.join(" ")).not.toMatch(/[0-9a-f]{16}/);
  });

  it("offers a generic safe user display helper for future app status surfaces", () => {
    expect(__test.safeUserDisplayName("slack-user-ref-db6b1d5cccad3cbf", { user_ref: "slack-user-ref-8d4e38d9a1cf7e80" } as any)).toBe("Cameron");
    expect(__test.safeUserDisplayName("slack-user-ref-8d4e38d9a1cf7e80", { user_ref: "slack-user-ref-db6b1d5cccad3cbf" } as any)).toBe("Avery");
    expect(__test.safeUserDisplayName("slack-user-ref-db6b1d5cccad3cbf", { user_ref: "slack-user-ref-db6b1d5cccad3cbf" } as any)).toBe("you");
    expect(__test.safeUserDisplayName("slack-user-ref-newteammate001122334455", { user_ref: "admin" } as any, { fallbackLabel: "a user" })).toBe("a user");
  });

  it("excludes an all-day event that only touches the window boundary at midnight", () => {
    // June 12 all-day Independence Day in Asia/Riyadh: starts Jun 12 00:00 +03, ends Jun 13 00:00 +03.
    const holiday = { start: { dateTime: "2026-06-12T00:00:00.0000000", timeZone: "Asia/Riyadh" }, end: { dateTime: "2026-06-13T00:00:00.0000000", timeZone: "Asia/Riyadh" } } as any;
    // Query window = Saturday June 13 in +03 (the holiday's end == window start).
    const satStart = "2026-06-12T21:00:00.000Z"; // 2026-06-13T00:00:00+03:00
    const satEnd = "2026-06-13T21:00:00.000Z";   // 2026-06-14T00:00:00+03:00
    expect(__test.eventOverlapsWindow(holiday, satStart, satEnd)).toBe(false);
    // The same holiday DOES overlap its own day (June 12 window).
    const friStart = "2026-06-11T21:00:00.000Z"; // 2026-06-12T00:00:00+03:00
    const friEnd = "2026-06-12T21:00:00.000Z";   // 2026-06-13T00:00:00+03:00
    expect(__test.eventOverlapsWindow(holiday, friStart, friEnd)).toBe(true);
    // A normal event inside the window stays.
    const meeting = { start: { dateTime: "2026-06-13T10:00:00.0000000", timeZone: "Asia/Riyadh" }, end: { dateTime: "2026-06-13T11:00:00.0000000", timeZone: "Asia/Riyadh" } } as any;
    expect(__test.eventOverlapsWindow(meeting, satStart, satEnd)).toBe(true);
    // An event starting exactly at the window end is excluded.
    const atEnd = { start: { dateTime: "2026-06-14T00:00:00.0000000", timeZone: "Asia/Riyadh" }, end: { dateTime: "2026-06-14T01:00:00.0000000", timeZone: "Asia/Riyadh" } } as any;
    expect(__test.eventOverlapsWindow(atEnd, satStart, satEnd)).toBe(false);
  });

  it("filters getSchedule team free/busy items against the same half-open window", () => {
    const window = {
      start: "2026-06-12T21:00:00.000Z",
      end: "2026-06-13T21:00:00.000Z",
      timezone: "Asia/Riyadh",
    };
    const targets = [{ user_ref: "member-a", schedule: "member-a", schedule_hash: "memberhash" }];
    const data = {
      value: [{
        scheduleItems: [
          {
            status: "busy",
            start: { dateTime: "2026-06-12T00:00:00.0000000", timeZone: "Asia/Riyadh" },
            end: { dateTime: "2026-06-13T00:00:00.0000000", timeZone: "Asia/Riyadh" },
          },
          {
            status: "busy",
            start: { dateTime: "2026-06-13T09:00:00.0000000", timeZone: "Asia/Riyadh" },
            end: { dateTime: "2026-06-13T10:00:00.0000000", timeZone: "Asia/Riyadh" },
          },
        ],
      }],
    };
    const events = __test.teamScheduleEventsInWindow(data, targets, window);
    expect(events).toHaveLength(1);
    expect(events[0].start).toMatchObject({ dateTime: "2026-06-13T09:00:00.0000000", timeZone: "Asia/Riyadh" });
  });

  it("maps Graph mailboxSettings timezones (Windows or IANA) to usable IANA ids", () => {
    expect(__test.ianaFromGraphTimeZone("Arabic Standard Time")).toBe("Asia/Baghdad");
    expect(__test.ianaFromGraphTimeZone("Turkey Standard Time")).toBe("Europe/Istanbul");
    expect(__test.ianaFromGraphTimeZone("Eastern Standard Time")).toBe("America/New_York");
    expect(__test.ianaFromGraphTimeZone("Asia/Baghdad")).toBe("Asia/Baghdad");
    expect(__test.ianaFromGraphTimeZone("Not A Zone")).toBeNull();
    expect(__test.ianaFromGraphTimeZone("")).toBeNull();
  });

  it("enumerates calendar groups and classifies own vs shared calendars", () => {
    expect(__test.graphCalendarGroupsListUrl()).toContain("/me/calendarGroups?");
    expect(__test.graphCalendarsInGroupUrl("grp 1")).toContain("/me/calendarGroups/grp%201/calendars?");
    expect(__test.calendarGroupIdsFromGraphData({ value: [{ id: "g1" }, { id: "g2" }, { id: "g1" }] })).toEqual(["g1", "g2"]);
    expect(__test.calendarDescriptorsFromGraphData({ value: [{ id: "c1" }, { id: "c2", isSharedWithMe: true }, { id: "c1" }] })).toEqual([
      { id: "c1", isSharedWithMe: false },
      { id: "c2", isSharedWithMe: true },
    ]);
  });
});
