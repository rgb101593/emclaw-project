// Canonical allowlists for scheduled, incremental, and alert operations.

export const allowedMemberCalendarOperations = new Set([
  "calendar.metadata_daily_status_check",
  "calendar.free_busy_availability_check",
  "calendar.conflict_check",
  "calendar.digest_metadata",
  "calendar.digest_metadata_delta",
  "calendar.daily_brief_metadata",
  "calendar.alert_metadata_delta",
  "calendar.important_meeting_alert_metadata",
  "calendar.delta_check",
]);

export const deltaCalendarOperations = new Set([
  "calendar.digest_metadata",
  "calendar.digest_metadata_delta",
  "calendar.daily_brief_metadata",
  "calendar.alert_metadata_delta",
  "calendar.important_meeting_alert_metadata",
  "calendar.delta_check",
]);

export const alertCalendarOperations = new Set([
  "calendar.alert_metadata_delta",
  "calendar.important_meeting_alert_metadata",
]);

export const allowedMemberMailOperations = new Set([
  "mail.alert_metadata_delta",
  "mail.important_email_alert_metadata",
  "mail.digest_metadata_delta",
]);

export const deltaMailOperations = new Set([
  "mail.alert_metadata_delta",
  "mail.important_email_alert_metadata",
  "mail.digest_metadata_delta",
]);

export const alertMailOperations = new Set([
  "mail.alert_metadata_delta",
  "mail.important_email_alert_metadata",
]);
