"""Unit tests for the macOS (AppleScript) server — no live Outlook required.

They cover two things live tests can't pin down reliably:
  1. Field parsing — the control-character delimiters, empty trailing fields,
     content containing the OLD "|||" / "===" delimiters.
  2. Script construction — that each tool emits the AppleScript we expect and
     rejects bad input before any script runs.
"""
import os
from datetime import datetime

import pytest

from outlook_desktop_mcp import server_mac as s
from outlook_desktop_mcp.applescript_bridge import AppleScriptError, AppleScriptTimeout
from outlook_desktop_mcp.utils.applescript_helpers import DELIM, RECORD_DELIM, date_var_lines
from tests.conftest import call, fields, rec


# --- input validation ---------------------------------------------------------

INJECTIONS = [
    "1\nend tell\ndo shell script \"touch /tmp/pwned\"",
    "1 or x",
    "-5",
    "+5",
    "1_0",
    "５",           # full-width digit: int() accepts it, AppleScript must not see it
    "",
    "999999999999",  # beyond AppleScript's integer range
]


@pytest.mark.parametrize("bad", INJECTIONS)
@pytest.mark.parametrize("tool", [
    s.read_email, s.mark_as_read, s.mark_as_unread, s.get_event,
    s.delete_event, s.get_task, s.complete_task, s.delete_task, s.list_attachments,
])
def test_entry_id_is_validated_before_any_script_runs(fake, tool, bad):
    out = call(tool(bad) if tool is not s.read_email else tool(entry_id=bad or "x"))
    assert out["code"] == "invalid_argument"
    assert fake.calls == 0


def test_ui_scraped_ids_get_a_clear_error(fake):
    out = call(s.read_email(entry_id="ui-3"))
    assert "UI-scraping" in out["error"] and fake.calls == 0


def test_valid_entry_id_is_interpolated_as_a_plain_integer(fake):
    fake.response = "Subject"
    call(s.mark_as_read(" 42 "))
    assert "set m to message id 42\n" in fake.last_script


@pytest.mark.parametrize("tool, kwargs", [
    (s.list_events, {}),
    (s.search_events, {"query": "x"}),
    (s.create_event, {"subject": "S", "start": "2026-09-10 10:00", "end": "2026-09-10 11:00"}),
    (s.create_meeting, {"subject": "S", "start": "2026-09-10 10:00", "end": "2026-09-10 11:00",
                        "required_attendees": "a@b.com"}),
])
def test_non_numeric_calendar_id_is_rejected(fake, tool, kwargs):
    out = call(tool(calendar_id='133; do shell script "boom"', **kwargs))
    assert "error" in out and fake.calls == 0


@pytest.mark.parametrize("tool, kwargs", [
    (s.list_events, {"start_date": "next tuesday"}),
    (s.search_events, {"query": "x", "end_date": "soon"}),
    (s.create_event, {"subject": "S", "start": "tomorrow", "end": "2026-09-10 11:00"}),
    (s.update_event, {"entry_id": "1", "start": "25/02/2026"}),
    (s.create_task, {"subject": "S", "due_date": "Friday"}),
    (s.update_task, {"entry_id": "1", "start_date": "?"}),
    (s.list_tasks, {"due_start": "last week"}),
    (s.snooze_email, {"entry_id": "42", "until": "next tuesday"}),
])
def test_bad_dates_return_json_errors_instead_of_raising(fake, tool, kwargs):
    out = call(tool(**kwargs))
    assert out["code"] == "invalid_argument" and "ISO 8601" in out["error"]
    assert fake.calls == 0


def test_timezone_aware_dates_are_converted_to_local_time(fake):
    fake.response = ""
    aware = "2026-09-10T08:00:00+00:00"
    local = datetime.fromisoformat(aware).astimezone().replace(tzinfo=None)
    call(s.list_events(start_date=aware, end_date="2026-09-11"))
    assert f"set time of startD to {local.hour * 3600 + local.minute * 60}" in fake.last_script


def test_bare_end_date_includes_the_whole_day(fake):
    fake.response = ""
    call(s.list_events(start_date="2026-09-10", end_date="2026-09-10"))
    assert f"set time of endD to {23 * 3600 + 59 * 60 + 59}" in fake.last_script


def test_count_is_clamped(fake):
    fake.response = ""
    call(s.search_emails("x", count=100000))
    assert "set maxCount to 200" in fake.last_script
    call(s.search_emails("x", count=-3))
    assert "set maxCount to 1" in fake.last_script


def test_empty_inbox_falls_back_to_ui_scraping(fake):
    fake.response = ""
    assert call(s.list_emails()) == []
    assert fake.calls == 2 and 'tell application "System Events"' in fake.last_script
    call(s.list_emails(folder="Archive"))
    assert fake.calls == 3


# --- error reporting ------------------------------------------------------------

def test_applescript_errors_carry_a_code(fake):
    fake.response = AppleScriptError("AppleScript error: Can't get message id 5. (-1728)", -1728)
    out = call(s.mark_as_read("5"))
    assert out["code"] == "not_found" and out["error"].startswith("Error marking email as read")


def test_timeout_on_a_mutating_tool_warns_the_action_may_have_happened(fake):
    fake.response = AppleScriptTimeout("AppleScript timed out after 30s")
    out = call(s.send_email("a@b.com", "S", "B"))
    assert out["code"] == "timeout" and "may still have completed" in out["error"]


def test_timeout_on_a_read_only_tool_has_no_such_warning(fake):
    fake.response = AppleScriptTimeout("AppleScript timed out after 30s")
    out = call(s.list_emails())
    assert out["code"] == "timeout" and "may still" not in out["error"]


# --- delimiters and empty trailing fields --------------------------------------

def test_list_tasks_preserves_content_containing_old_delimiters(fake):
    tricky = "Report |||draft=== v2"
    fake.response = rec("101", tricky, "", "not completed", "priority high")
    out = call(s.list_tasks())
    assert out == [{"entry_id": "101", "subject": tricky, "due_date": None,
                    "complete": False, "priority": "priority high"}]


def test_get_task_parses_a_task_without_start_date(fake):
    # The last field is empty. Stripping the output used to drop it together
    # with its delimiter, and the tool failed with "Failed to parse task data".
    fake.response = fields("55", "My Task", "", "completed", "priority normal",
                           "notes with ||| and === inside", "")
    out = call(s.get_task("55"))
    assert out["body"] == "notes with ||| and === inside"
    assert out["complete"] is True and out["start_date"] is None


def test_read_email_parses_an_empty_body(fake):
    fake.response = fields("9", "Hi", "a@b.com", "A", "2026-09-10T10:00:00", "true", "0",
                           "me@x.com; ", "", "")
    out = call(s.read_email(entry_id="9"))
    assert out["body"] == "" and out["to"] == "me@x.com" and out["cc"] == ""
    assert out["unread"] is False


def test_list_calendars_keeps_a_calendar_with_an_empty_trailing_count(fake):
    fake.response = rec("133", "Docencia", "jagalindo@us.es", "") + rec("5", "Other", "", "7")
    cals = call(s.list_calendars())
    assert [c["calendar_id"] for c in cals] == ["133", "5"]
    assert cals[0]["event_count"] is None


def test_search_tasks_returns_all_matching_records(fake):
    fake.response = (rec("1", "Call plumber", "", "not completed", "priority low")
                     + rec("2", "Call bank", "", "not completed", "priority high"))
    out = call(s.search_tasks("call"))
    assert len(out) == 2
    assert 'tasks whose name contains "call"' in fake.last_script


# --- dates come back as ISO 8601 ----------------------------------------------

@pytest.mark.parametrize("tool, kwargs, expected", [
    (s.list_emails, {}, "my isoDate(time received of m)"),
    (s.search_emails, {"query": "x"}, "my isoDate(time received of m)"),
    (s.list_tasks, {}, "my isoDate(due date of t)"),
    (s.search_tasks, {"query": "x"}, "my isoDate(due date of t)"),
])
def test_dates_are_rendered_without_locale_coercion(fake, tool, kwargs, expected):
    fake.response = ""
    call(tool(**kwargs))
    script = fake.scripts[0]  # list_emails may add a UI-scraping fallback call
    assert expected in script and "as string" not in script


def test_emails_read_sender_through_its_record(fake):
    # "address of sender of m" fails in Outlook; the record must be fetched first.
    fake.response = ""
    call(s.list_emails())
    assert "set senderRec to sender of m" in fake.scripts[0]
    assert "(address of sender of" not in fake.scripts[0]


# --- email tools ------------------------------------------------------------------

def test_send_email_returns_json(fake):
    out = call(s.send_email("a@b.com; c@d.com", "Hello", "Line 1\nLine 2", cc="e@f.com"))
    assert out["status"] == "sent"
    scr = fake.last_script
    assert 'address:"a@b.com"' in scr and 'address:"c@d.com"' in scr
    assert "make new cc recipient" in scr and "send newMsg" in scr
    assert "Line 1<br>Line 2" in scr


def test_create_draft_persists_without_send_or_save(fake):
    fake.response = fields("990", "Hello")
    out = call(s.create_draft("a@b.com", "Hello", "hi"))
    assert out["entry_id"] == "990"
    assert "save newMsg" not in fake.last_script and "send newMsg" not in fake.last_script
    assert 'address:"a@b.com"' in fake.last_script


def test_send_email_rejects_missing_attachment(fake):
    out = call(s.send_email("a@b.com", "S", "B", attachments=["/no/such/file.pdf"]))
    assert "not found" in out["error"].lower() and fake.calls == 0


def test_send_email_attaches_a_real_file(fake):
    call(s.send_email("a@b.com", "S", "B", attachments=[os.path.abspath(__file__)]))
    assert "make new attachment at newMsg" in fake.last_script
    assert "POSIX file" in fake.last_script


def test_read_email_by_subject_reports_not_found(fake):
    fake.response = "NOT_FOUND"
    out = call(s.read_email(subject_search="zzz"))
    assert out["code"] == "not_found"
    assert 'whose subject contains "zzz"' in fake.last_script


def test_read_email_requires_an_id_or_subject(fake):
    out = call(s.read_email())
    assert out["code"] == "invalid_argument" and fake.calls == 0


def test_move_email_returns_the_new_id(fake):
    fake.response = fields("501", "Invoice")
    out = call(s.move_email("42", "Archive"))
    assert out == {"status": "moved", "entry_id": "501", "previous_entry_id": "42",
                   "subject": "Invoice", "folder": "Archive"}
    assert 'set m to move m to mail folder "Archive"' in fake.last_script


def test_mark_as_read_returns_json(fake):
    fake.response = "Hello"
    assert call(s.mark_as_read("5")) == {"status": "read", "entry_id": "5", "subject": "Hello"}


def test_snooze_email_flags_and_sets_reminder(fake):
    fake.response = fields("42", "Some subject")
    out = call(s.snooze_email("42", "2026-07-25 09:00"))
    scr = fake.last_script
    assert "set todo flag of m to not completed" in scr
    assert "set reminder date time of m to remD" in scr and "set year of remD to 2026" in scr
    assert 'date "' not in scr and "set due date of m to remD" in scr
    assert "move m to" not in scr
    assert out["reminder"] == "2026-07-25T09:00:00"


def test_snooze_email_with_folder_moves_and_returns_new_id(fake):
    fake.response = fields("77", "Some subject")
    out = call(s.snooze_email("42", "2026-07-25T09:00:00", move_to_folder="Pospuesto"))
    assert 'set m to move m to mail folder "Pospuesto"' in fake.last_script
    assert out["entry_id"] == "77"


def test_unsnooze_email_clears_flag_and_reminder(fake):
    fake.response = fields("42", "Some subject")
    call(s.unsnooze_email("42"))
    scr = fake.last_script
    assert "set todo flag of m to not flagged" in scr
    assert "set reminder date time of m to missing value" in scr
    assert "set due date of m to missing value" in scr and "move m to" not in scr

    fake.response = fields("88", "Some subject")
    out = call(s.unsnooze_email("42", move_to_inbox=True))
    assert "set m to move m to inbox" in fake.last_script and out["entry_id"] == "88"


def test_reply_email_returns_json(fake):
    fake.response = "Original"
    out = call(s.reply_email("42", "Thanks", reply_all=True))
    assert out["status"] == "sent" and out["reply_all"] is True
    assert "set replyMsg to reply all to m" in fake.last_script


def test_forward_email_uses_forward_verb(fake):
    fake.response = "Original subject"
    out = call(s.forward_email("42", "x@y.com", comment="FYI"))
    scr = fake.last_script
    assert "set fwdMsg to forward m" in scr and 'address:"x@y.com"' in scr
    assert "content of fwdMsg" in scr
    assert out["status"] == "forwarded"


def test_list_folders_walks_to_max_depth(fake):
    fake.response = (rec("7", "Inbox", "Inbox", "0", "Work", "10", "2")
                     + rec("9", "Projects", "Inbox/Projects", "1", "Work", "3", "0"))
    out = call(s.list_folders(max_depth=3))
    assert out[1] == {"folder_id": "9", "name": "Projects", "path": "Inbox/Projects",
                      "depth": 1, "account": "Work", "item_count": 3, "unread_count": 0}
    assert 'my walkFolder(contents of f, 0, 3, "")' in fake.last_script


# --- calendars and events ---------------------------------------------------------

def evrec(eid, subject, start, end, cal_id="132", cal="Diverso", rec_flag="false"):
    return rec(eid, subject, start, end, "missing value", "jagalindo@us.es",
               "false", cal_id, cal, rec_flag)


def test_list_calendars_parses_records(fake):
    fake.response = (rec("133", "Docencia", "jagalindo@us.es", "212")
                     + rec("205", "malawito@gmail.com", "", "2330"))
    cals = call(s.list_calendars())
    assert cals[0] == {"calendar_id": "133", "name": "Docencia",
                       "account": "jagalindo@us.es", "event_count": 212}
    assert cals[1]["account"] == "" and cals[1]["event_count"] == 2330
    assert "if cname is not missing value then" in fake.last_script


def test_list_events_scopes_and_filters_in_applescript(fake):
    fake.response = evrec("1", "A", "2026-09-10T10:00:00", "2026-09-10T11:00:00")
    call(s.list_events(start_date="2026-09-01", end_date="2026-09-30", calendar_id="133"))
    scr = fake.last_script
    assert "calendar events of calendar id 133 whose" in scr
    assert "start time is greater than or equal to startD" in scr
    assert "start time is less than or equal to endD" in scr
    assert "set year of startD to 2026" in scr and 'date "' not in scr
    assert "on isoDate(d)" in scr and "my isoDate(start time of e)" in scr

    call(s.list_events(start_date="2026-09-01", end_date="2026-09-30"))
    assert "set evts to (calendar events whose" in fake.last_script


def test_list_events_sorts_then_applies_count(fake):
    fake.response = (evrec("3", "C", "2026-09-30T09:00:00", "2026-09-30T10:00:00")
                     + evrec("1", "A", "2026-09-10T09:00:00", "2026-09-10T10:00:00")
                     + evrec("2", "B", "2026-09-20T09:00:00", "2026-09-20T10:00:00"))
    evs = call(s.list_events(start_date="2026-09-01", end_date="2026-09-30"))
    assert [e["subject"] for e in evs] == ["A", "B", "C"]
    evs = call(s.list_events(start_date="2026-09-01", end_date="2026-09-30", count=2))
    assert [e["subject"] for e in evs] == ["A", "B"]


def test_list_events_reports_calendar_and_recurrence(fake):
    fake.response = evrec("7", "Clase", "2026-09-10T10:00:00", "2026-09-10T11:00:00",
                          cal_id="133", cal="Docencia", rec_flag="true")
    ev = call(s.list_events(start_date="2026-09-01", end_date="2026-09-30"))[0]
    assert ev["calendar_id"] == "133" and ev["calendar"] == "Docencia"
    assert ev["is_recurring"] is True and ev["location"] == ""


def test_search_events_combines_subject_and_dates(fake):
    fake.response = evrec("1", "Reunión", "2026-09-10T10:00:00", "2026-09-10T11:00:00")
    call(s.search_events(query="reuni", start_date="2026-09-01", end_date="2026-09-30",
                         calendar_id="132"))
    scr = fake.last_script
    assert 'calendar events of calendar id 132 whose subject contains "reuni"' in scr
    assert "start time is greater than or equal to startD" in scr


def test_get_event_reports_calendar_attendees_and_recurrence(fake):
    fake.response = fields("7", "Clase", "2026-09-10T10:00:00", "2026-09-10T11:00:00",
                           "Aula F0.10", "jagalindo@us.es", "false", "cuerpo",
                           "a@b.com; c@d.com; ", "133", "Docencia", "true")
    ev = call(s.get_event("7"))
    assert ev["calendar_id"] == "133" and ev["calendar"] == "Docencia"
    assert ev["is_recurring"] is True and ev["start"] == "2026-09-10T10:00:00"
    assert ev["attendees"] == "a@b.com; c@d.com"
    assert "set ea to email address of a" in fake.last_script


def test_create_event_uses_date_vars_and_reminder(fake):
    fake.response = fields("99", "E", "2026-07-03T00:00:00", "2026-07-04T00:00:00", "131", "Cal")
    call(s.create_event(subject="E", start="2026-07-03", end="2026-07-04", all_day=True))
    scr = fake.last_script
    assert "start time:startD" in scr and "end time:endD" in scr
    assert 'start time:date "' not in scr
    assert "set year of startD to 2026" in scr and "set year of endD to 2026" in scr
    assert "has reminder:true, reminder time:15" in scr

    call(s.create_event(subject="E", start="2026-07-03", end="2026-07-04", reminder_minutes=0))
    assert "has reminder:false" in fake.last_script


def test_create_event_targets_the_requested_calendar(fake):
    fake.response = fields("99", "S", "2026-09-10T10:00:00", "2026-09-10T11:00:00",
                           "135", "Tutorías")
    out = call(s.create_event(subject="S", start="2026-09-10 10:00",
                              end="2026-09-10 11:00", calendar_id="135"))
    assert "make new calendar event at calendar id 135 with properties" in fake.last_script
    assert out["calendar_id"] == "135" and out["calendar"] == "Tutorías"

    call(s.create_event(subject="S", start="2026-09-10 10:00", end="2026-09-10 11:00"))
    assert "make new calendar event with properties" in fake.last_script


def test_create_meeting_adds_attendees_and_sends(fake):
    fake.response = fields("99", "M", "2026-09-10T10:00:00", "2026-09-10T11:00:00", "131", "Cal")
    out = call(s.create_meeting("M", "2026-09-10 10:00", "2026-09-10 11:00",
                                "a@b.com; c@d.com", optional_attendees="e@f.com"))
    scr = fake.last_script
    assert 'make new required attendee at newEvt with properties {email address:{address:"a@b.com"}}' in scr
    assert 'make new optional attendee at newEvt with properties {email address:{address:"e@f.com"}}' in scr
    assert "send meeting newEvt" in scr
    assert scr.index("make new required attendee") < scr.index("send meeting newEvt")
    assert out["status"] == "sent" and out["entry_id"] == "99"


def test_create_meeting_can_skip_sending(fake):
    fake.response = fields("99", "M", "2026-09-10T10:00:00", "2026-09-10T11:00:00", "131", "Cal")
    out = call(s.create_meeting("M", "2026-09-10 10:00", "2026-09-10 11:00", "a@b.com",
                                send_invites=False))
    assert "send meeting" not in fake.last_script and out["status"] == "created"


def test_create_meeting_requires_an_attendee(fake):
    out = call(s.create_meeting("M", "2026-09-10 10:00", "2026-09-10 11:00", " ; "))
    assert out["code"] == "invalid_argument" and fake.calls == 0


def test_update_event_sets_all_day_before_times(fake):
    fake.response = fields("5", "X", "2026-09-10T10:00:00", "2026-09-10T11:00:00", "", "false")
    out = call(s.update_event("5", start="2026-09-10 10:00", all_day=False))
    scr = fake.last_script
    assert scr.index("set all day flag of e to false") < scr.index("set start time of e to startD")
    assert out["all_day"] is False and out["location"] == ""


def test_update_event_guards_its_result_fields(fake):
    fake.response = fields("5", "X", "", "", "", "")
    call(s.update_event("5", subject="X"))
    assert "set eloc to (location of e) as text" in fake.last_script


def test_update_event_without_fields_is_rejected(fake):
    out = call(s.update_event("5"))
    assert out["code"] == "invalid_argument" and fake.calls == 0


def test_delete_event_returns_json(fake):
    fake.response = "Standup"
    assert call(s.delete_event("8")) == {"status": "deleted", "entry_id": "8", "subject": "Standup"}


# --- move_event -----------------------------------------------------------------------

def test_move_event_refuses_an_event_with_attendees(fake):
    fake.response = fields("REFUSED", "3", "Reunión con el comité", "133")
    out = call(s.move_event("14000", "135"))
    assert out["status"] == "refused" and out["attendees"] == 3
    scr = fake.last_script
    assert "count of (attendees of srcEv)" in scr
    assert scr.index("if nAtt > 0 then") < scr.index("duplicate srcEv")


def test_move_event_only_deletes_after_verifying_the_copy(fake):
    fake.response = fields("REFUSED", "3", "x", "133")
    call(s.move_event("14000", "135"))
    scr = fake.last_script
    delete_at = scr.index("delete (calendar event id")
    assert scr.index("if copySubject is not srcSubject then") < delete_at
    assert scr.index("if copyId is 0 then") < delete_at
    assert "delete (calendar event id 14000 of calendar id srcCal)" in scr


def test_move_event_bails_out_without_a_source_calendar(fake):
    fake.response = fields("NOSRCCAL", "0", "Clase", "-1")
    out = call(s.move_event("14000", "135"))
    assert out["status"] == "failed" and "nothing was changed" in out["reason"]
    scr = fake.last_script
    assert scr.index("if srcCal is -1 then") < scr.index("duplicate srcEv")


def test_move_event_reports_the_new_id(fake):
    fake.response = fields("MOVED", "14915", "Clase IISSI2", "133")
    out = call(s.move_event("14000", "135"))
    assert out["status"] == "moved" and out["entry_id"] == "14915"
    assert out["previous_entry_id"] == "14000" and out["to_calendar_id"] == "135"


@pytest.mark.parametrize("status, expected", [
    ("NOOP", "unchanged"),
    ("NOCOPY", "failed"),
    ("MISMATCH", "failed"),
    ("NOSTART", "failed"),
    ("NODUPRECUR", "failed"),
])
def test_move_event_outcomes(fake, status, expected):
    fake.response = fields(status, "0", "Clase", "133")
    assert call(s.move_event("14000", "135"))["status"] == expected


@pytest.mark.parametrize("kwargs", [
    {"entry_id": "14000", "target_calendar_id": "135; delete"},
    {"entry_id": "no", "target_calendar_id": "135"},
    {"entry_id": "14000", "target_calendar_id": ""},
])
def test_move_event_rejects_bad_ids(fake, kwargs):
    out = call(s.move_event(**kwargs))
    assert "error" in out and fake.calls == 0


# --- respond_to_meeting ---------------------------------------------------------------

def test_respond_accept_with_comment(fake):
    fake.response = "Team sync"
    out = call(s.respond_to_meeting("71", "accept", comment="See you there"))
    scr = fake.last_script
    assert "accept invite mm" in scr and "set mm to meeting message id 71" in scr
    assert "sending response true" in scr and 'comment "See you there"' in scr
    assert out["status"] == "responded" and out["response"] == "accept"


def test_respond_tentative_and_decline(fake):
    fake.response = "Team sync"
    call(s.respond_to_meeting("71", "tentative"))
    assert "accept tentatively invite mm" in fake.last_script
    call(s.respond_to_meeting("71", "decline", send_response=False, comment="skip"))
    scr = fake.last_script
    assert "decline invite mm" in scr
    assert "sending response false" in scr and "comment" not in scr


def test_respond_rejects_invalid_response(fake):
    out = call(s.respond_to_meeting("71", "maybe"))
    assert "error" in out and fake.calls == 0


# --- tasks --------------------------------------------------------------------------------

def test_list_tasks_filters_due_dates_in_applescript(fake):
    fake.response = ""
    call(s.list_tasks(due_start="2026-03-10"))
    scr = fake.last_script
    assert "due date is greater than or equal to dueStartD" in scr
    assert "set day of dueStartD to 10" in scr

    call(s.list_tasks(due_end="2026-03-20"))
    scr = fake.last_script
    assert "due date is less than or equal to dueEndD" in scr
    assert f"set time of dueEndD to {23 * 3600 + 59 * 60 + 59}" in scr

    call(s.list_tasks(due_start="2026-03-10", due_end="2026-03-20", include_completed=True))
    assert ("(tasks whose due date is greater than or equal to dueStartD and "
            "due date is less than or equal to dueEndD)") in fake.last_script


def test_list_tasks_without_filter_keeps_undated_tasks(fake):
    fake.response = (rec("1", "Early", "2026-03-01T09:00:00", "not completed", "priority normal")
                     + rec("4", "NoDue", "", "not completed", "priority normal"))
    out = call(s.list_tasks())
    assert [t["due_date"] for t in out] == ["2026-03-01T09:00:00", None]
    assert "(tasks whose todo flag is not completed)" in fake.last_script


def test_update_task_builds_set_lines(fake):
    fake.response = fields("7", "Renamed", "", "not completed", "priority high")
    out = call(s.update_task("7", subject="Renamed", importance="high", complete=False))
    scr = fake.last_script
    assert 'set name of t to "Renamed"' in scr
    assert "set priority of t to priority high" in scr
    assert "set todo flag of t to not completed" in scr
    assert out["status"] == "updated" and out["priority"] == "priority high"


def test_update_task_rejects_empty_or_invalid_input(fake):
    assert "error" in call(s.update_task("7"))
    assert "error" in call(s.update_task("7", importance="urgent"))
    assert fake.calls == 0


def test_create_task_rejects_invalid_importance(fake):
    # It used to fall back to "normal" silently, unlike update_task.
    out = call(s.create_task("x", importance="urgent"))
    assert out["code"] == "invalid_argument" and fake.calls == 0


def test_create_task_returns_the_record(fake):
    fake.response = fields("12", "Buy milk", "2026-03-01T00:00:00", "not completed", "priority low")
    out = call(s.create_task("Buy milk", due_date="2026-03-01", importance="low"))
    assert out["status"] == "created" and out["entry_id"] == "12"
    assert out["due_date"] == "2026-03-01T00:00:00"
    assert "priority:priority low" in fake.last_script


def test_complete_and_delete_task_return_json(fake):
    fake.response = "Buy milk"
    assert call(s.complete_task("3"))["status"] == "completed"
    assert call(s.delete_task("3"))["status"] == "deleted"


# --- categories ---------------------------------------------------------------------------

def test_list_categories(fake):
    fake.response = rec("Work", "1") + rec("Personal", "2")
    assert [c["name"] for c in call(s.list_categories())] == ["Work", "Personal"]


@pytest.mark.parametrize("item_type, ref", [
    ("task", "task id 77"), ("email", "message id 77"), ("event", "calendar event id 77"),
])
def test_set_category_targets_the_item_type(fake, item_type, ref):
    fake.response = "Item"
    out = call(s.set_category("77", "Work, Follow-up", item_type=item_type))
    scr = fake.last_script
    assert f"set theItem to {ref}" in scr
    assert '"Work"' in scr and '"Follow-up"' in scr
    assert "make new category with properties" in scr
    assert out["categories"] == ["Work", "Follow-up"]


def test_set_category_clear_and_invalid_type(fake):
    fake.response = "Cleared"
    call(s.set_category("3", "", item_type="task"))
    assert "repeat with nm in {}" in fake.last_script
    n = fake.calls
    assert "error" in call(s.set_category("1", "X", item_type="bogus"))
    assert fake.calls == n


# --- attachments --------------------------------------------------------------------------

def test_list_attachments(fake):
    fake.response = rec("1", "a.pdf", "1200") + rec("2", "b.png", "")
    out = call(s.list_attachments("5"))
    assert out == [{"index": 1, "filename": "a.pdf", "size": 1200},
                   {"index": 2, "filename": "b.png", "size": 0}]


@pytest.mark.parametrize("name, expected", [
    ("report.pdf", "report.pdf"),
    ("../../.zshrc", ".zshrc"),
    ("..\\..\\evil.bat", "evil.bat"),
    ("/etc/passwd", "passwd"),
    ("..", "attachment-2"),
    ("", "attachment-2"),
])
def test_safe_filename(name, expected):
    assert s._safe_filename(name, 2) == expected


def test_save_attachment_sanitizes_and_never_overwrites(fake, tmp_path):
    (tmp_path / ".zshrc").write_text("mine")
    fake.response = lambda script: "OK" + DELIM + "../../.zshrc" if "NOATT" in script else ""
    out = call(s.save_attachment("5", 1, str(tmp_path)))
    assert out["path"] == str(tmp_path / ".zshrc (1)")
    assert f'save a in (POSIX file "{tmp_path}/.zshrc (1)")' in fake.last_script
    assert (tmp_path / ".zshrc").read_text() == "mine"


def test_save_attachment_reports_a_missing_index(fake, tmp_path):
    fake.response = "NOATT" + DELIM + "1"
    out = call(s.save_attachment("5", 3, str(tmp_path)))
    assert out["code"] == "not_found" and fake.calls == 1


def test_save_attachment_rejects_index_zero(fake, tmp_path):
    out = call(s.save_attachment("5", 0, str(tmp_path)))
    assert out["code"] == "invalid_argument" and fake.calls == 0


# --- helpers ------------------------------------------------------------------------------

def test_parse_records_keeps_empty_trailing_fields():
    raw = DELIM.join(["1", "a", ""]) + RECORD_DELIM + DELIM.join(["2", "", ""]) + RECORD_DELIM
    assert s._parse_records(raw, 3) == [["1", "a", ""], ["2", "", ""]]


def test_parse_records_skips_short_and_blank_records():
    raw = "1" + DELIM + "a" + RECORD_DELIM + "\n" + RECORD_DELIM
    assert s._parse_records(raw, 3) == []


def test_split_fields_lets_the_last_field_hold_the_rest():
    assert s._split_fields("a" + DELIM + "b" + DELIM + "c", 2) == ["a", "b" + DELIM + "c"]
    assert s._split_fields("a", 2) is None


def test_date_var_lines_is_locale_independent():
    lines = date_var_lines("d", datetime(2026, 7, 3, 0, 0, 0))
    assert 'date "' not in lines and "set year of d to 2026" in lines
    assert lines.index("set month of d to 1") < lines.index("set month of d to 7")
    assert lines.index("set day of d to 1") < lines.index("set day of d to 3")
    tlines = date_var_lines("d", datetime(2026, 7, 3, 14, 30, 15))
    assert f"set time of d to {14 * 3600 + 30 * 60 + 15}" in tlines


def test_every_tool_is_registered_with_its_parameters():
    import asyncio
    tools = {t.name: t for t in asyncio.run(s.mcp.list_tools())}
    assert len(tools) == 34
    assert "send_invites" in tools["create_meeting"].inputSchema["properties"]
    assert tools["get_event"].description.startswith("Read the full details")
