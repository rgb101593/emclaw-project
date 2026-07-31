import { createHash } from "node:crypto";

import { ROSTER_PATH } from "./config.js";
import { normalizeUserRef, readJsonFile } from "./util.js";
import type { Caller, RosterUser } from "./types.js";

export function loadRoster(): RosterUser[] {
  const data = readJsonFile(ROSTER_PATH);
  if (!data || typeof data !== "object" || !Array.isArray((data as { users?: unknown }).users)) {
    throw new Error("roster_schema_users_array_required");
  }
  return ((data as { users: unknown[] }).users).map((item) => {
    const user = item as Record<string, unknown>;
    if (typeof user.user_ref !== "string" || typeof user.role !== "string" || typeof user.status !== "string") {
      throw new Error("invalid_roster_user_entry");
    }
    return {
      user_ref: user.user_ref,
      role: user.role,
      status: user.status,
      display_style: typeof user.display_style === "string" ? user.display_style : undefined,
      slack_user_sha256: typeof user.slack_user_sha256 === "string" ? user.slack_user_sha256 : undefined,
      slack_user_hash_short: typeof user.slack_user_hash_short === "string" ? user.slack_user_hash_short : undefined,
    };
  });
}

export function resolveCaller(callerUserRef: unknown): Caller {
  const userRef = normalizeUserRef(callerUserRef);
  const user = loadRoster().find((entry) => entry.user_ref === userRef);
  if (!user) throw new Error("caller_not_found_in_roster");
  if (user.status !== "active") throw new Error("caller_not_active");
  return { user_ref: user.user_ref, role: user.role, status: user.status };
}

export function requireAdmin(caller: Caller): void {
  if (caller.role !== "admin") throw new Error("admin_role_required");
}

export function requireSelfCapable(caller: Caller): void {
  if (caller.role !== "admin" && caller.role !== "member") throw new Error("member_or_admin_role_required");
}

export function sha256Hex(value: string): string {
  return createHash("sha256").update(value).digest("hex");
}

export function hashCandidates(value: unknown): string[] {
  if (typeof value !== "string" || value.trim() === "") return [];
  const raw = value.trim();
  const variants = Array.from(new Set([raw, raw.toUpperCase(), raw.toLowerCase()]));
  return variants.map((candidate) => sha256Hex(candidate));
}
