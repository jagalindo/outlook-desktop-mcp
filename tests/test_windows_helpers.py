"""Tests for the Windows server's pure helpers. They avoid importing
server.py, which needs pywin32, so they run on any OS."""
from datetime import datetime, timezone

import pytest

from outlook_desktop_mcp.utils.com_dates import parse_date, task_reminder_time
from outlook_desktop_mcp.utils.errors import format_com_error


def test_parse_date_returns_a_datetime_not_a_string():
    # COM would parse a raw string with the Windows locale ("2026-03-04" can
    # become 3 April), so properties must be given a datetime.
    assert parse_date(" 2026-03-04 14:00 ") == datetime(2026, 3, 4, 14, 0)


def test_parse_date_rejects_non_iso_input():
    with pytest.raises(ValueError, match="ISO 8601"):
        parse_date("04/03/2026")


def test_task_reminder_time_counts_back_from_the_due_date():
    assert task_reminder_time(datetime(2026, 3, 4, 9, 0), 30) == datetime(2026, 3, 4, 8, 30)


def test_task_reminder_time_drops_pywin32_tzinfo():
    # pywin32 labels Outlook's local wall-clock times as UTC.
    due = datetime(2026, 3, 4, 9, 0, tzinfo=timezone.utc)
    assert task_reminder_time(due, 60) == datetime(2026, 3, 4, 8, 0)


@pytest.mark.parametrize("due", [None, datetime(4501, 1, 1)])
def test_task_reminder_time_needs_a_due_date(due):
    with pytest.raises(ValueError, match="due date"):
        task_reminder_time(due, 15)


def test_format_com_error_passes_our_own_messages_through():
    msg = "Account 'x' not found. Use list_accounts to see available accounts."
    assert format_com_error(ValueError(msg)) == msg
    assert format_com_error(FileNotFoundError("Attachment not found: /a.pdf")) == "Attachment not found: /a.pdf"
    assert format_com_error(RuntimeError()) == "RuntimeError"
