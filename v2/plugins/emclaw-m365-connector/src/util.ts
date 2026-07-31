import { mkdirSync, readFileSync, renameSync, statSync, writeFileSync } from "node:fs";
import { TOKEN_DIR } from "./config.js";
import type { ToolResult } from "./types.js";

export function normalizeUserRef(value: unknown): string {
  if (typeof value !== "string" || value.trim() === "") return "admin";
  return value.trim();
}

export function safeFileNamePart(value: string): string {
  if (!/^[A-Za-z0-9._-]{1,120}$/.test(value)) throw new Error("invalid_safe_file_name");
  return value;
}

export function tokenFileNameFor(userRef: string): string {
  return userRef === "admin" ? "admin-token.json" : `${userRef}-token.json`;
}

export function tokenFileExistsByStat(userRef: string): boolean {
  try {
    return statSync(`${TOKEN_DIR}/${tokenFileNameFor(userRef)}`).isFile();
  } catch {
    return false;
  }
}

// Redact credential and identity patterns before values are logged or returned.
export function redact(value: unknown): unknown {
  if (typeof value === "string") {
    return value
      .replace(/xox[baprs]-[A-Za-z0-9-]+/g, "[REDACTED_SLACK_TOKEN]")
      .replace(/eyJ[A-Za-z0-9_-]{20,}\.[A-Za-z0-9_-]{20,}\.[A-Za-z0-9_-]{10,}/g, "[REDACTED_JWT]")
      .replace(/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/g, "[REDACTED_GUID]")
      .replace(/\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/g, "[REDACTED_EMAIL]")
      .replace(/\b[UW][A-Z0-9]{8,}\b/g, "[REDACTED_SLACK_ID]");
  }
  if (Array.isArray(value)) return value.map((item) => redact(item));
  if (value && typeof value === "object") {
    const out: Record<string, unknown> = {};
    for (const [key, item] of Object.entries(value as Record<string, unknown>)) out[key] = redact(item);
    return out;
  }
  return value;
}

export function okResult(extra: ToolResult): ToolResult {
  return {
    ok: true,
    connector: "microsoft365",
    workflow_tools_step: "step_3_connector_wrappers",
    graph_called: false,
    oauth_performed: false,
    token_contents_read: false,
    m365_content_accessed: false,
    slack_send_performed: false,
    real_execution_implemented: false,
    ...extra,
  };
}

export function errorResult(code: string, message: string, extra: ToolResult = {}): ToolResult {
  return okResult({ ok: false, error_code: code, message: redact(message), ...redact(extra) as ToolResult });
}

export function readJsonFile(path: string): unknown {
  return JSON.parse(readFileSync(path, "utf8"));
}

export function writeJsonAtomic(path: string, data: unknown): void {
  const tmp = `${path}.tmp-${Date.now()}`;
  writeFileSync(tmp, `${JSON.stringify(data, null, 2)}\n`, { mode: 0o644 });
  renameSync(tmp, path);
}

export function ensureDir(path: string): void {
  mkdirSync(path, { recursive: true, mode: 0o755 });
}

export function writePrivateJsonAtomic(path: string, data: unknown): void {
  const tmp = `${path}.tmp-${Date.now()}`;
  writeFileSync(tmp, `${JSON.stringify(data, null, 2)}\n`, { mode: 0o600 });
  renameSync(tmp, path);
}

export function isoNow(): string {
  return new Date().toISOString();
}
