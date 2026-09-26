"""Date helpers for the Windows (COM) server.

Kept free of pywin32 so they can be unit-tested on any OS.
"""
from datetime import datetime, timedelta


def parse_date(date_str: str) -> datetime:
    """Parse ISO 8601 date string like '2026-02-25 14:00' or '2026-02-25T14:00:00'.

    Assign the result to COM date properties rather than the raw string: COM
    converts a string using the Windows locale, so "2026-03-04" can land on
    the 3rd of April.
    """
    try:
        return datetime.fromisoformat(date_str.strip())
    except ValueError:
        raise ValueError(
            f"Invalid date {date_str!r}. Expected ISO 8601, e.g. \"2026-02-25 14:00\"."
        ) from None


# Outlook reports "no date" as 1 January 4501.
_NO_DATE_YEAR = 4501


def task_reminder_time(due, minutes: int) -> datetime:
    """The absolute reminder time for a task due at `due`.

    Tasks have no ReminderMinutesBeforeStart (that is an appointment
    property), so the reminder is set through ReminderTime instead.
    """
    if due is None or due.year >= _NO_DATE_YEAR:
        raise ValueError("reminder_minutes needs a due date on the task")
    return due.replace(tzinfo=None) - timedelta(minutes=minutes)
