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

    # --- 4. create_draft: persists without send/save and parses id -----------
    # `make new outgoing message` already persists to Drafts; an explicit
    # `save` fails with -1701 on outgoing messages (requires an `in <file>`).
    fake.response = "990" + DELIM + "Hello"
    out = json.loads(await s.create_draft("a@b.com", "Hello", "hi"))
    check("create_draft parses entry_id", out.get("entry_id") == "990", detail=out)
    check("create_draft does not `send` or `save`",
          "save newMsg" not in fake.last_script and "send newMsg" not in fake.last_script)
    check("create_draft adds recipient", 'address:"a@b.com"' in fake.last_script)

    # --- 4b. snooze_email / unsnooze_email -----------------------------------
    fake.response = "42" + DELIM + "Some subject"
    out = json.loads(await s.snooze_email("42", "2026-07-25 09:00"))
    scr = fake.last_script
    check("snooze_email flags for follow-up", "set todo flag of m to not completed" in scr)
    check("snooze_email sets reminder from date vars",
          "set reminder date time of m to remD" in scr and "set year of remD to 2026" in scr
          and 'date "' not in scr)
    check("snooze_email sets due date", "set due date of m to remD" in scr)
    check("snooze_email without folder does not move", "move m to" not in scr)
    check("snooze_email returns reminder ISO", out.get("reminder") == "2026-07-25T09:00:00",
          detail=out)

    fake.response = "77" + DELIM + "Some subject"
    out = json.loads(await s.snooze_email("42", "2026-07-25T09:00:00", move_to_folder="Pospuesto"))
    scr = fake.last_script
    check("snooze_email with folder moves the message",
          'set m to move m to mail folder "Pospuesto"' in scr)
    check("snooze_email returns post-move entry_id", out.get("entry_id") == "77", detail=out)

    fake.calls = 0
    out = json.loads(await s.snooze_email("42", "next tuesday"))
    check("snooze_email rejects invalid datetime without running script",
          "error" in out and fake.calls == 0, detail=out)

    fake.response = "42" + DELIM + "Some subject"
    out = json.loads(await s.unsnooze_email("42"))
    scr = fake.last_script
    check("unsnooze_email clears flag", "set todo flag of m to not flagged" in scr)
    check("unsnooze_email clears reminder and dates",
          "set reminder date time of m to missing value" in scr
          and "set due date of m to missing value" in scr)
    check("unsnooze_email without move keeps folder", "move m to" not in scr)

    fake.response = "88" + DELIM + "Some subject"
    out = json.loads(await s.unsnooze_email("42", move_to_inbox=True))
    check("unsnooze_email move_to_inbox moves to inbox",
          "set m to move m to inbox" in fake.last_script)
    check("unsnooze_email returns post-move entry_id", out.get("entry_id") == "88", detail=out)

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

    # The due filter runs inside Outlook's `whose` clause, so it sees every task
    # rather than only the first few fetched.
    fake.response = ""
    await s.list_tasks(due_start="2026-03-10")
    scr = fake.last_script
    check("list_tasks due_start filters in AppleScript",
          "due date is greater than or equal to dueStartD" in scr
          and "set day of dueStartD to 10" in scr, detail=scr)

    await s.list_tasks(due_end="2026-03-20")
    scr = fake.last_script
    check("list_tasks due_end is inclusive of the whole day",
          "due date is less than or equal to dueEndD" in scr
          and f"set time of dueEndD to {23*3600 + 59*60 + 59}" in scr, detail=scr)

    await s.list_tasks(due_start="2026-03-10", due_end="2026-03-20", include_completed=True)
    scr = fake.last_script
    check("list_tasks due range combines both bounds without the completion filter",
          "(tasks whose due date is greater than or equal to dueStartD and due date is less than or equal to dueEndD)" in scr,
          detail=scr)

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

    # --- 12. calendars -------------------------------------------------------
    # list_calendars: parse id/name/account/count, including a calendar Outlook
    # attributes to no account (shared/subscribed ones report an empty account).
    fake.response = (rec("133", "Docencia", "jagalindo@us.es", "212")
                     + rec("205", "malawito@gmail.com", "", "2330"))
    cals = json.loads(await s.list_calendars())
    check("list_calendars parses id/name/account/count",
          len(cals) == 2 and cals[0]["calendar_id"] == "133"
          and cals[0]["name"] == "Docencia"
          and cals[0]["account"] == "jagalindo@us.es"
          and cals[0]["event_count"] == 212, detail=f"got {cals}")
    check("list_calendars leaves an unattributed calendar's account empty",
          cals[1]["account"] == "" and cals[1]["event_count"] == 2330, detail=f"got {cals}")
    check("list_calendars skips the nameless account roots",
          "if cname is not missing value then" in fake.last_script)

    # A full event record as the AppleScript side builds it.
    def evrec(eid, subject, start, end, cal_id="132", cal="Diverso", rec_flag="false"):
        return rec(eid, subject, start, end, "missing value", "jagalindo@us.es",
                   "false", cal_id, cal, rec_flag)

    # list_events: scoping to one calendar
    fake.response = evrec("1", "A", "2026-09-10T10:00:00", "2026-09-10T11:00:00")
    await s.list_events(start_date="2026-09-01", end_date="2026-09-30", calendar_id="133")
    scr = fake.last_script
    check("list_events with calendar_id scopes to that calendar",
          "calendar events of calendar id 133 whose" in scr)
    check("list_events filters by date in AppleScript, not by overfetching",
          "start time is greater than or equal to startD" in scr
          and "start time is less than or equal to endD" in scr)
    check("list_events builds bounds with locale-independent date vars",
          "set year of startD to 2026" in scr and 'date "' not in scr)
    check("list_events emits the ISO date handlers",
          "on isoDate(d)" in scr and "my isoDate(start time of e)" in scr)

    # list_events: no calendar_id spans every calendar
    await s.list_events(start_date="2026-09-01", end_date="2026-09-30")
    check("list_events without calendar_id spans all calendars",
          "set evts to (calendar events whose" in fake.last_script)

    # list_events: sorting by start time and the count limit
    fake.response = (evrec("3", "C", "2026-09-30T09:00:00", "2026-09-30T10:00:00")
                     + evrec("1", "A", "2026-09-10T09:00:00", "2026-09-10T10:00:00")
                     + evrec("2", "B", "2026-09-20T09:00:00", "2026-09-20T10:00:00"))
    evs = json.loads(await s.list_events(start_date="2026-09-01", end_date="2026-09-30"))
    check("list_events sorts by start time",
          [e["subject"] for e in evs] == ["A", "B", "C"], detail=f"got {evs}")
    evs = json.loads(await s.list_events(start_date="2026-09-01", end_date="2026-09-30", count=2))
    check("list_events applies count after sorting, keeping the earliest",
          [e["subject"] for e in evs] == ["A", "B"], detail=f"got {evs}")

    # list_events: the new per-event fields
    fake.response = evrec("7", "Clase", "2026-09-10T10:00:00", "2026-09-10T11:00:00",
                          cal_id="133", cal="Docencia", rec_flag="true")
    evs = json.loads(await s.list_events(start_date="2026-09-01", end_date="2026-09-30"))
    check("list_events reports each event's calendar and recurrence",
          evs[0]["calendar_id"] == "133" and evs[0]["calendar"] == "Docencia"
          and evs[0]["is_recurring"] is True, detail=f"got {evs}")
    check("list_events blanks a 'missing value' location",
          evs[0]["location"] == "", detail=f"got {evs}")

    # search_events: subject and date clauses, scoped
    fake.response = evrec("1", "Reunión", "2026-09-10T10:00:00", "2026-09-10T11:00:00")
    await s.search_events(query="reuni", start_date="2026-09-01", end_date="2026-09-30",
                          calendar_id="132")
    scr = fake.last_script
    check("search_events combines subject and date filters on one calendar",
          'calendar events of calendar id 132 whose subject contains "reuni"' in scr
          and "start time is greater than or equal to startD" in scr)

    # invalid calendar_id is rejected before any script runs
    for tool, kwargs in (
        (s.list_events, {}),
        (s.search_events, {"query": "x"}),
        (s.create_event, {"subject": "S", "start": "2026-09-10 10:00", "end": "2026-09-10 11:00"}),
    ):
        fake.calls = 0
        out = json.loads(await tool(calendar_id="133; do shell script \"boom\"", **kwargs))
        check(f"{tool.__name__} rejects a non-numeric calendar_id without running a script",
              "error" in out and fake.calls == 0, detail=f"got {out}")

    # create_event: target calendar
    fake.response = DELIM.join(["99", "S", "2026-09-10T10:00:00", "2026-09-10T11:00:00",
                                "135", "Tutorías"])
    out = json.loads(await s.create_event(subject="S", start="2026-09-10 10:00",
                                          end="2026-09-10 11:00", calendar_id="135"))
    check("create_event targets the requested calendar",
          "make new calendar event at calendar id 135 with properties" in fake.last_script)
    check("create_event reports the calendar the event landed in",
          out["calendar_id"] == "135" and out["calendar"] == "Tutorías", detail=f"got {out}")
    await s.create_event(subject="S", start="2026-09-10 10:00", end="2026-09-10 11:00")
    check("create_event without calendar_id uses the default calendar",
          "make new calendar event with properties" in fake.last_script)

    # get_event: calendar and recurrence
    fake.response = DELIM.join(["7", "Clase", "2026-09-10T10:00:00", "2026-09-10T11:00:00",
                                "Aula F0.10", "jagalindo@us.es", "false", "cuerpo",
                                "a@b.com; ", "133", "Docencia", "true"])
    got = json.loads(await s.get_event("7"))
    check("get_event reports the calendar and recurrence",
          got["calendar_id"] == "133" and got["calendar"] == "Docencia"
          and got["is_recurring"] is True and got["start"] == "2026-09-10T10:00:00",
          detail=f"got {got}")

    # --- 13. move_event ------------------------------------------------------
    # The guard: an event with attendees must never be copied or deleted.
    fake.response = DELIM.join(["REFUSED", "3", "Reunión con el comité", "133"])
    out = json.loads(await s.move_event("14000", "135"))
    check("move_event refuses an event that has attendees",
          out["status"] == "refused" and out["attendees"] == 3, detail=f"got {out}")
    check("move_event asks Outlook for the attendee count before anything else",
          "count of (attendees of srcEv)" in fake.last_script)
    check("move_event returns before duplicating when there are attendees",
          fake.last_script.index("if nAtt > 0 then")
          < fake.last_script.index("duplicate srcEv"))

    # The original is only deleted after the copy is found and verified.
    scr = fake.last_script
    check("move_event verifies the copy's subject before deleting the original",
          scr.index("if copySubject is not srcSubject then") < scr.index("delete (calendar event id"))
    check("move_event bails out when no copy appeared",
          scr.index("if copyId is 0 then") < scr.index("delete (calendar event id"))
    check("move_event deletes the original by explicit id, not a loop reference",
          "delete (calendar event id 14000 of calendar id srcCal)" in scr)

    fake.response = DELIM.join(["MOVED", "14915", "Clase IISSI2", "133"])
    out = json.loads(await s.move_event("14000", "135"))
    check("move_event reports the new entry_id after a move",
          out["status"] == "moved" and out["entry_id"] == "14915"
          and out["previous_entry_id"] == "14000"
          and out["to_calendar_id"] == "135", detail=f"got {out}")

    fake.response = DELIM.join(["NOOP", "0", "Clase", "135"])
    out = json.loads(await s.move_event("14000", "135"))
    check("move_event is a no-op when the event is already in the target",
          out["status"] == "unchanged", detail=f"got {out}")

    fake.response = DELIM.join(["NOCOPY", "0", "Clase", "133"])
    out = json.loads(await s.move_event("14000", "135"))
    check("move_event reports failure, not success, when the copy never appeared",
          out["status"] == "failed" and "untouched" in out["reason"], detail=f"got {out}")

    fake.response = DELIM.join(["MISMATCH", "99", "Otro evento", "133"])
    out = json.loads(await s.move_event("14000", "135"))
    check("move_event reports failure when the copy does not match the original",
          out["status"] == "failed", detail=f"got {out}")

    for bad_kwargs in ({"entry_id": "14000", "target_calendar_id": "135; delete"},
                       {"entry_id": "no", "target_calendar_id": "135"}):
        fake.calls = 0
        out = json.loads(await s.move_event(**bad_kwargs))
        check(f"move_event rejects {bad_kwargs} without running a script",
              "error" in out and fake.calls == 0, detail=f"got {out}")

    print(f"\n{passed}/{total} unit checks passed")
    return 0 if passed == total else 1


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
