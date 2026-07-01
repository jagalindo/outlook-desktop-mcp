"""
Unit tests for the macOS (AppleScript) server — NO live Outlook required.

Unlike the integration tests in this folder, these tests replace the
AppleScript bridge with a fake, so they run anywhere with just the `mcp`
dependency installed:

    python tests/unit_mac_test.py

They focus on two things the integration tests can't cover reliably:
  1. Field parsing — especially the control-character delimiters, proving that
     content containing the OLD "|||" / "===" delimiters is parsed correctly.
  2. Script construction — that each tool emits the AppleScript we expect
     (recipients, attachments, update `set` lines, forward verb, etc.).
"""
import sys
import os
import json
import asyncio

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp import server_mac as s
from outlook_desktop_mcp.utils.applescript_helpers import DELIM, RECORD_DELIM

passed = 0
total = 0


def check(name, cond, detail=""):
    global passed, total
    total += 1
    if cond:
        passed += 1
        print(f"  PASS: {name}")
    else:
        print(f"  FAIL: {name} {detail}")


class FakeBridge:
    """Stand-in for AppleScriptBridge. Records the last script it was asked to
    run and returns a programmable response (str or callable)."""

    def __init__(self):
        self.last_script = None
        self.calls = 0
        self.response = ""

    async def run(self, script, timeout=None):
        self.last_script = script
        self.calls += 1
        if callable(self.response):
            return self.response(script)
        return self.response


def rec(*fields):
    """Join fields with the field delimiter and terminate with a record delim,
    exactly as the AppleScript side builds list output."""
    return DELIM.join(fields) + RECORD_DELIM


async def main():
    fake = FakeBridge()
    s.bridge = fake  # tools reference the module-global `bridge`

    # --- 1. Delimiter robustness: content containing old "|||"/"===" ---------
    tricky_name = "Report |||draft=== v2"
    fake.response = rec("101", tricky_name, "missing value", "not completed", "priority high")
    out = json.loads(await s.list_tasks())
    check(
        "list_tasks preserves content containing old delimiters",
        len(out) == 1 and out[0]["subject"] == tricky_name,
        detail=f"got {out}",
    )
    check("list_tasks maps completion flag", out[0]["complete"] is False)

    # --- 2. get_task parsing with delimiter-laden body -----------------------
    body = "notes with ||| and === inside"
    fake.response = DELIM.join(
        ["55", "My Task", "missing value", "completed", "priority normal", body, "missing value"]
    )
    out = json.loads(await s.get_task("55"))
    check("get_task parses body with embedded old delimiters", out["body"] == body, detail=out)
    check("get_task maps completed flag", out["complete"] is True)

    # --- 3. search_tasks: parses multiple records ----------------------------
    fake.response = (
        rec("1", "Call plumber", "missing value", "not completed", "priority low")
        + rec("2", "Call bank", "missing value", "not completed", "priority high")
    )
    out = json.loads(await s.search_tasks("call"))
    check("search_tasks returns all matching records", len(out) == 2, detail=out)
    check("search_tasks escapes query into `contains`", 'contains "call"' in fake.last_script)

    # --- 4. create_draft: saves (does not send) and parses id ----------------
    fake.response = "990" + DELIM + "Hello"
    out = json.loads(await s.create_draft("a@b.com", "Hello", "hi"))
    check("create_draft parses entry_id", out.get("entry_id") == "990", detail=out)
    check("create_draft uses `save` not `send`",
          "save newMsg" in fake.last_script and "send newMsg" not in fake.last_script)

    # --- 5. update_task: builds the right `set` lines ------------------------
    fake.response = rec("7", "Renamed", "not completed", "priority high").rstrip(RECORD_DELIM)
    await s.update_task("7", subject="Renamed", importance="high", complete=False)
    scr = fake.last_script
    check("update_task sets name", 'set name of t to "Renamed"' in scr)
    check("update_task sets priority", "set priority of t to priority high" in scr)
    check("update_task reopens task", "set todo flag of t to not completed" in scr)

    # update_task with no fields must not call the bridge
    fake.calls = 0
    out = json.loads(await s.update_task("7"))
    check("update_task with no fields returns error, no bridge call",
          out.get("error") and fake.calls == 0, detail=out)

    # invalid importance is rejected before running any script
    fake.calls = 0
    out = json.loads(await s.update_task("7", importance="urgent"))
    check("update_task rejects invalid importance",
          "error" in out and fake.calls == 0, detail=out)

    # --- 6. forward_email: uses the forward verb + recipients ----------------
    fake.response = "Original subject"
    res = await s.forward_email("42", "x@y.com", comment="FYI")
    scr = fake.last_script
    check("forward_email uses `forward` verb", "set fwdMsg to forward m" in scr)
    check("forward_email adds recipient", 'address:"x@y.com"' in scr)
    check("forward_email prepends comment", 'content of fwdMsg' in scr)
    check("forward_email confirms", "forwarded" in res.lower())

    # --- 7. attachment validation happens before running any script ----------
    fake.calls = 0
    out = json.loads(await s.send_email("a@b.com", "S", "B", attachments=["/no/such/file.pdf"]))
    check("send_email rejects missing attachment without running script",
          "error" in out and "not found" in out["error"].lower() and fake.calls == 0,
          detail=out)

    # valid attachment path produces a `make new attachment` line
    fake.calls = 0
    fake.response = ""
    here = os.path.abspath(__file__)
    await s.send_email("a@b.com", "S", "B", attachments=[here])
    check("send_email emits attachment line for a real file",
          "make new attachment at newMsg" in fake.last_script and "POSIX file" in fake.last_script)

    # --- 8. list_tasks due-date filtering ------------------------------------
    def _tasks_payload():
        return (
            rec("1", "Early", "2026-03-01 09:00:00", "not completed", "priority normal")
            + rec("2", "Mid", "2026-03-15 09:00:00", "not completed", "priority normal")
            + rec("3", "Late", "2026-04-01 09:00:00", "not completed", "priority normal")
            + rec("4", "NoDue", "missing value", "not completed", "priority normal")
        )

    fake.response = _tasks_payload()
    out = json.loads(await s.list_tasks(due_start="2026-03-10"))
    subs = [t["subject"] for t in out]
    check("list_tasks due_start excludes earlier and undated tasks",
          subs == ["Mid", "Late"], detail=subs)

    fake.response = _tasks_payload()
    out = json.loads(await s.list_tasks(due_end="2026-03-20"))
    subs = [t["subject"] for t in out]
    check("list_tasks due_end excludes later and undated tasks",
          subs == ["Early", "Mid"], detail=subs)

    fake.response = _tasks_payload()
    out = json.loads(await s.list_tasks(due_start="2026-03-10", due_end="2026-03-20"))
    subs = [t["subject"] for t in out]
    check("list_tasks due range keeps only in-window task", subs == ["Mid"], detail=subs)

    fake.response = _tasks_payload()
    out = json.loads(await s.list_tasks())
    check("list_tasks without filter keeps all (incl. undated)", len(out) == 4, detail=out)

    print(f"\n{passed}/{total} unit checks passed")
    return 0 if passed == total else 1


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
