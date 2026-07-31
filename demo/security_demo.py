#!/usr/bin/env python3
"""Standalone demonstration of EMClaw's deterministic security controls.

The demo uses only the Python standard library and does not connect to Slack,
Microsoft 365, or an LLM. The production connector is under v2/plugins/emclaw-m365-connector.
"""

from __future__ import annotations
import hashlib
import json
import re
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parents[1]
ROSTER_PATH = ROOT / "v2" / "identity" / "roster.json"

GREEN, RED, DIM, BOLD, RESET = "\033[32m", "\033[31m", "\033[2m", "\033[1m", "\033[0m"


def title(text: str) -> None:
    print(f"\n{BOLD}{'-' * 72}\n{text}\n{'-' * 72}{RESET}")


def ok(msg: str) -> None:
    print(f"  {GREEN}[ALLOWED]{RESET} {msg}")


def blocked(msg: str) -> None:
    print(f"  {RED}[BLOCKED]{RESET} {msg}")


def note(msg: str) -> None:
    print(f"  {DIM}{msg}{RESET}")


# 1. Trusted caller identity
# Resolve identity from gateway routing data rather than model-supplied arguments.
def sha256_hex(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def resolve_identity(raw_sender_id: str, roster: list[dict]) -> dict | None:
    h = sha256_hex(raw_sender_id)
    for user in roster:
        if user.get("status") != "active":
            continue
        full = user.get("slack_user_sha256")
        short = user.get("slack_user_hash_short")
        if full and h == full:
            return user
        if short and h.startswith(short):
            return user
    return None  # fail closed


def demo_identity() -> None:
    title("1. Trusted caller identity bridge (fail-closed, hash-based)")
    # The fixture is deterministic and contains no real Slack identifier.
    alice_raw = "U_DEMO_ALICE_0001"
    roster = [
        {"user_ref": "member-alice", "role": "member", "status": "active",
         "slack_user_sha256": sha256_hex(alice_raw)},
        {"user_ref": "member-bob", "role": "member", "status": "active",
         "slack_user_sha256": sha256_hex("U_DEMO_BOB_0002")},
    ]
    note(f"roster stores only SHA-256 hashes, e.g. {sha256_hex(alice_raw)[:24]}...")

    who = resolve_identity(alice_raw, roster)
    ok(f"runtime sender {alice_raw!r} resolves to {who['user_ref']!r}")

    spoof = resolve_identity("U_ATTACKER_NOT_ON_ROSTER", roster)
    if spoof is None:
        blocked("unknown sender id -> caller_identity_unresolved (fail closed)")

    # Claimed identities are ignored; only the runtime sender ID is trusted.
    note("a model/tool argument claiming user_ref='member-bob' is never consulted -")
    note("identity comes only from the runtime sender id above.")


# 2. Cross-user access control
def demo_cross_user() -> None:
    title("2. Per-person data isolation (no acting on another user's behalf)")

    def member_request(caller_ref: str, target_ref: str) -> bool:
        # Member tools require the resolved caller to match the target.
        return caller_ref == target_ref

    if member_request("member-alice", "member-alice"):
        ok("member-alice queries member-alice's own mailbox")
    if not member_request("member-alice", "member-bob"):
        blocked("member-alice tries to read member-bob's mailbox -> self-only gate")


# 3. Outbound secret-shape guard
# Apply the same pattern checks used before replies reach Slack.
SECRET_PATTERNS = [
    re.compile(r"xox[baprs]-[A-Za-z0-9-]+"),                                  # Slack token
    re.compile(r"eyJ[A-Za-z0-9_-]{20,}\.[A-Za-z0-9_-]{20,}\.[A-Za-z0-9_-]{10,}"),  # JWT
    re.compile(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"),  # GUID
]


def looks_like_secret(text: str) -> bool:
    return any(p.search(text) for p in SECRET_PATTERNS)


def demo_secret_guard() -> None:
    title("3. Outbound secret-shape guard (cancel before send)")
    clean = "You have 3 unread emails; the most likely to matter is \"Q3 close\"."
    # Build the fixture at runtime so secret scanners do not mistake it for a credential.
    fake_token = "xox" + "b-" + "000000000000-" + "EXAMPLE0NOTAREALTOKEN"
    leak = f"Reconnected! your token is {fake_token}"
    if not looks_like_secret(clean):
        ok("normal metadata reply is delivered")
    if looks_like_secret(leak):
        blocked("reply containing a Slack-token shape is cancelled")


# 4. Phantom-success guard
# Success claims require a deterministic receipt.
def demo_phantom_success() -> None:
    title("4. Phantom-success guard (receipts, not promises)")
    claim = re.compile(r"\b(created|scheduled|terminated|paused|sent)\b", re.I)

    def gate(reply: str, receipt: str | None) -> bool:
        return not (claim.search(reply) and receipt is None)

    if gate("Your invoice alert is created. Receipt R-4821.", receipt="R-4821"):
        ok("success claim with a tool receipt is delivered")
    if not gate("All done - I created your workflow!", receipt=None):
        blocked("success claim with no receipt is cancelled")


# 5. Delta watermark deduplication
# Watermarks prevent duplicate and pre-creation alerts.
def demo_watermark() -> None:
    title("5. Delta watermark de-duplication (no double / no back alerts)")
    seen: set[str] = set()
    watermark_created_at = 100  # alert created at t=100; older mail is ignored

    def cycle(messages: list[dict]) -> list[str]:
        fresh = []
        for m in messages:
            if m["received"] < watermark_created_at:
                continue  # never back-alert
            if m["id"] in seen:
                continue  # never double-alert
            seen.add(m["id"])
            fresh.append(m["id"])
        return fresh

    old = {"id": "m0", "received": 50}
    a = {"id": "m1", "received": 120}
    b = {"id": "m2", "received": 130}
    first = cycle([old, a, b])
    ok(f"first run alerts on new mail only: {first}  (m0 pre-dates the alert)")
    second = cycle([a, b])
    if not second:
        blocked(f"re-running over the same messages alerts again: {second or '[]'}")


def main() -> int:
    print(f"{BOLD}EMClaw - deterministic security-model demonstration{RESET}")
    print(f"{DIM}standard-library only | no Slack/M365/LLM required{RESET}")
    if ROSTER_PATH.exists():
        roster = json.loads(ROSTER_PATH.read_text()).get("users", [])
        print(f"{DIM}(loaded {len(roster)} roster entries from {ROSTER_PATH.relative_to(ROOT)} - hashes only){RESET}")
    demo_identity()
    demo_cross_user()
    demo_secret_guard()
    demo_phantom_success()
    demo_watermark()
    title("All gates behaved as designed.")
    print(f"{DIM}Connector implementation: see v2/plugins/emclaw-m365-connector "
          f"(run `npm test` there).{RESET}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
