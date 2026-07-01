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
from outlook_desktop_mcp.utils.applescript_helpers import DELIM, RECORD_DELIM, date_var_lines
from datetime import datetime

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

    # --- 9. categories -------------------------------------------------------
    fake.response = rec("Work", "1") + rec("Personal", "2")
    out = json.loads(await s.list_categories())
    check("list_categories parses name/color records",
          [c["name"] for c in out] == ["Work", "Personal"], detail=out)

    # set_category targets the item type explicitly (no ambiguous resolution)
    fake.response = "My Task"
    await s.set_category("77", "Work, Follow-up", item_type="task")
    scr = fake.last_script
    check("set_category task uses `task id`", "set theItem to task id 77" in scr)
    check("set_category builds requested name list",
          '"Work"' in scr and '"Follow-up"' in scr)
    check("set_category auto-creates missing categories",
          "make new category with properties" in scr)

    fake.response = "Some Email"
    await s.set_category("5", "Important", item_type="email")
    check("set_category email uses `message id`", "set theItem to message id 5" in fake.last_script)

    fake.response = "Some Event"
    await s.set_category("9", "Personal", item_type="event")
    check("set_category event uses `calendar event id`",
          "set theItem to calendar event id 9" in fake.last_script)

    fake.calls = 0
    out = json.loads(await s.set_category("1", "X", item_type="bogus"))
    check("set_category rejects invalid item_type without running script",
          "error" in out and fake.calls == 0, detail=out)

    # clearing categories yields an empty AppleScript list literal
    fake.calls = 0
    fake.response = "Cleared Task"
    await s.set_category("3", "", item_type="task")
    check("set_category with empty string builds empty list `{}`",
          "repeat with nm in {}" in fake.last_script)

    # --- 10. respond_to_meeting ---------------------------------------------
    fake.response = "Team sync"
    await s.respond_to_meeting("71", "accept", comment="See you there")
    scr = fake.last_script
    check("respond accept uses `accept invite` on a meeting message",
          "accept invite mm" in scr and "set mm to meeting message id 71" in scr)
    check("respond accept sends response with comment",
          "sending response true" in scr and 'comment "See you there"' in scr)

    fake.response = "Team sync"
    await s.respond_to_meeting("71", "tentative")
    check("respond tentative uses tentative verb",
          "accept tentatively invite mm" in fake.last_script)

    fake.response = "Team sync"
    await s.respond_to_meeting("71", "decline", send_response=False, comment="skip")
    scr = fake.last_script
    check("respond decline uses decline verb", "decline invite mm" in scr)
    check("respond with send_response=false omits response and comment",
          "sending response false" in scr and "comment" not in scr)

    fake.calls = 0
    out = json.loads(await s.respond_to_meeting("71", "maybe"))
    check("respond rejects invalid response without running script",
          "error" in out and fake.calls == 0, detail=out)

    # --- 11. locale-independent date construction ----------------------------
    lines = date_var_lines("d", datetime(2026, 7, 3, 0, 0, 0))
    check("date_var_lines builds from numeric components (no `date \"...\"`)",
          'date "' not in lines and "set year of d to 2026" in lines)
    check("date_var_lines resets month/day to 1 before applying real values",
          lines.index("set month of d to 1") < lines.index("set month of d to 7")
          and lines.index("set day of d to 1") < lines.index("set day of d to 3"))
    tlines = date_var_lines("d", datetime(2026, 7, 3, 14, 30, 15))
    check("date_var_lines encodes time as seconds since midnight",
          f"set time of d to {14*3600 + 30*60 + 15}" in tlines)

    # create_event must use the date variables, not a locale-parsed date string
    fake.response = "99" + DELIM + "E" + DELIM + "s" + DELIM + "e"
    await s.create_event(subject="E", start="2026-07-03", end="2026-07-04", all_day=True)
    scr = fake.last_script
    check("create_event uses date vars in properties, not date \"...\"",
          "start time:startD" in scr and "end time:endD" in scr and 'start time:date "' not in scr)
    check("create_event emits the date-builder lines",
          "set year of startD to 2026" in scr and "set year of endD to 2026" in scr)

    # --- entry_id validation: AppleScript injection must be rejected ---------
    import tempfile

    evil_id = '1\nend tell\ntell application "Finder" to activate\ntell application "Microsoft Outlook"'
    calls_before = fake.calls
    res = await s.mark_as_read(evil_id)
    check("mark_as_read rejects non-numeric entry_id without running a script",
          res.startswith("Error") and fake.calls == calls_before, detail=res)
    res = await s.read_email(entry_id=evil_id)
    check("read_email rejects non-numeric entry_id", res.startswith("Error") and fake.calls == calls_before)
    res = await s.delete_event(entry_id="42; delete every calendar event")
    check("delete_event rejects non-numeric entry_id", res.startswith("Error") and fake.calls == calls_before)
    res = await s.set_category(entry_id="x", categories="Work", item_type="task")
    check("set_category rejects non-numeric entry_id", res.startswith("Error") and fake.calls == calls_before)

    fake.response = "Some subject"
    res = await s.mark_as_read(" 042 ")
    check("mark_as_read accepts numeric entry_id and normalizes it",
          "message id 42" in fake.last_script and "Marked as read" in res, detail=fake.last_script)

    # --- count clamping -------------------------------------------------------
    fake.response = ""
    await s.search_emails(query="x", folder="drafts", count=99999)
    check("search_emails clamps count to 200", "set maxCount to 200" in fake.last_script)
    await s.search_emails(query="x", folder="drafts", count=0)
    check("search_emails clamps count up to 1", "set maxCount to 1" in fake.last_script)
    await s.list_tasks(count=99999)
    check("list_tasks clamps count to 200", "set maxCount to 200" in fake.last_script)

    # --- save_attachment: hostile filename cannot escape save_directory ------
    tmpdir = tempfile.mkdtemp(prefix="odm-test-")
    fake.response = lambda script: (
        "../../../etc/evil.txt" if "return name of a" in script else "OK"
    )
    out = json.loads(await s.save_attachment(entry_id="7", save_directory=tmpdir))
    check("save_attachment strips traversal components from filename",
          out["filename"] == "evil.txt", detail=out)
    check("save_attachment keeps path inside save_directory",
          os.path.realpath(out["path"]).startswith(os.path.realpath(tmpdir) + os.sep), detail=out)
    check("save_attachment saves via a Python-built path (no aname concatenation)",
          'save a in "' in fake.last_script and "& aname" not in fake.last_script)

    calls_before = fake.calls
    res = await s.save_attachment(entry_id="bad; rm", save_directory=tmpdir)
    check("save_attachment rejects non-numeric entry_id",
          res.startswith("Error") and fake.calls == calls_before)

    # --- list_events filters to the requested date range ---------------------
    fake.response = (
        rec("1", "InRange", "2026-03-22 14:00:00", "2026-03-22 15:00:00", "", "", "false")
        + rec("2", "TooEarly", "2026-03-01 09:00:00", "2026-03-01 10:00:00", "", "", "false")
        + rec("3", "TooLate", "2026-05-01 09:00:00", "2026-05-01 10:00:00", "", "", "false")
        + rec("4", "Unparseable", "someday maybe", "", "", "", "false")
    )
    out = json.loads(await s.list_events(start_date="2026-03-20", end_date="2026-03-25"))
    subs = [e["subject"] for e in out]
    check("list_events keeps only events inside the range (unparseable kept)",
          subs == ["InRange", "Unparseable"], detail=subs)

    # --- list_folders honors max_depth ---------------------------------------
    fake.response = rec("Inbox", "5", "2", "1") + rec("Sub", "1", "0", "2")
    out = json.loads(await s.list_folders(max_depth=2))
    scr = fake.last_script
    check("list_folders depth 2 recurses one level",
          "repeat with f1 in mail folders" in scr
          and "repeat with f2 in mail folders of f1" in scr
          and "f3" not in scr, detail=scr)
    check("list_folders parses depth field",
          out[0]["depth"] == 1 and out[1]["depth"] == 2, detail=out)
    await s.list_folders(max_depth=1)
    check("list_folders depth 1 stays top-level", "mail folders of" not in fake.last_script)
    await s.list_folders(max_depth=99)
    check("list_folders clamps depth to 3",
          "mail folders of f2" in fake.last_script and "f4" not in fake.last_script)

    print(f"\n{passed}/{total} unit checks passed")
    return 0 if passed == total else 1


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
