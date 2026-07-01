"""Helpers for building and parsing AppleScript safely."""
import re
from datetime import datetime


def escape(text: str) -> str:
    """Escape a string for safe embedding inside AppleScript double quotes.

    Handles backslashes, double quotes, and other special characters.
    """
    text = text.replace("\\", "\\\\")
    text = text.replace('"', '\\"')
    text = text.replace("\n", "\\n")
    text = text.replace("\r", "\\r")
    text = text.replace("\t", "\\t")
    return text


def safe_id(entry_id) -> str:
    """Validate an entry_id as a plain integer and return it as a string.

    macOS Outlook item IDs are numeric. Because IDs are interpolated
    unquoted into AppleScript (``message id 42``), anything non-numeric
    could inject arbitrary script — so reject it here.

    Raises ValueError if entry_id is not a non-negative integer.
    """
    text = str(entry_id).strip()
    if not text.isdigit():
        raise ValueError(
            f"entry_id must be numeric (got {text!r}). "
            "Use the id returned by list/search tools."
        )
    return str(int(text))  # normalize (strips leading zeros/unicode digits)


def format_date(dt: datetime) -> str:
    """Convert a Python datetime to an AppleScript date string.

    Returns a string like: date "Sunday, March 22, 2026 at 2:00:00 PM"
    AppleScript parses dates based on the system locale, so we use a
    locale-friendly format that osascript can interpret.

    DEPRECATED for building dates to assign to Outlook items: AppleScript's
    ``date "..."`` coercion is locale-dependent and misreads this ISO string
    on non-US systems (e.g. a Spanish-locale Mac parses "2026-07-03" into a
    wildly wrong date). Use ``date_var_lines`` instead, which builds the date
    from numeric components. This is kept only for backwards compatibility.
    """
    return f'date "{dt.strftime("%Y-%m-%d %H:%M:%S")}"'


def date_var_lines(var: str, dt: datetime) -> str:
    """Return AppleScript statements assigning a locale-independent date to `var`.

    The date is built from numeric components rather than via ``date "..."``
    string coercion, which AppleScript parses using the system's locale and
    therefore misreads ISO date strings on non-US systems. Month and day are
    reset to 1 before the real values are applied, so intermediate assignments
    can never overflow a short month (e.g. setting month to February while the
    day is still 31).

    Example output (var="d")::

        set d to (current date)
        set year of d to 2026
        set month of d to 1
        set day of d to 1
        set month of d to 7
        set day of d to 3
        set time of d to 0
    """
    secs = dt.hour * 3600 + dt.minute * 60 + dt.second
    return (
        f'set {var} to (current date)\n'
        f'set year of {var} to {dt.year}\n'
        f'set month of {var} to 1\n'
        f'set day of {var} to 1\n'
        f'set month of {var} to {dt.month}\n'
        f'set day of {var} to {dt.day}\n'
        f'set time of {var} to {secs}\n'
    )


def parse_date(text: str) -> str:
    """Parse an AppleScript date string to ISO 8601 format.

    AppleScript dates look like: "Sunday, March 22, 2026 at 2:00:00 PM"
    or various locale-specific formats. We attempt several common patterns.
    """
    text = text.strip()
    # Remove day name prefix if present (e.g., "Sunday, ")
    text = re.sub(r"^\w+day,\s*", "", text)
    # Remove " at " between date and time
    text = text.replace(" at ", " ")
    # Try common formats
    for fmt in (
        "%B %d, %Y %I:%M:%S %p",   # March 22, 2026 2:00:00 PM
        "%d. %B %Y %H:%M:%S",       # 22. mars 2026 14:00:00 (Norwegian)
        "%Y-%m-%d %H:%M:%S",        # 2026-03-22 14:00:00
        "%d/%m/%Y %H:%M:%S",        # 22/03/2026 14:00:00
        "%m/%d/%Y %H:%M:%S",        # 03/22/2026 14:00:00
    ):
        try:
            dt = datetime.strptime(text, fmt)
            return dt.isoformat()
        except ValueError:
            continue
    # Fallback: return as-is
    return text


# Locale-independent AppleScript folder keywords
FOLDER_MAP = {
    "inbox": "inbox",
    "sent": "sent items",
    "sentmail": "sent items",
    "sent items": "sent items",
    "drafts": "drafts",
    "deleted": "deleted items",
    "deleted items": "deleted items",
    "trash": "deleted items",
    "junk": "junk mail",
    "spam": "junk mail",
    "outbox": "outbox",
}


def resolve_folder_ref(folder_name: str) -> str:
    """Map a user-facing folder name to an AppleScript folder reference.

    Returns an AppleScript expression like 'inbox' or 'mail folder "Archive"'.
    Built-in folders use locale-independent keywords; custom folders use name lookup.
    """
    key = folder_name.lower().strip()
    if key in FOLDER_MAP:
        return FOLDER_MAP[key]
    # Custom folder — search by name
    return f'mail folder "{escape(folder_name)}"'


# Delimiters used for structured AppleScript output.
# These are ASCII control characters (Unit Separator / Record Separator) that
# effectively never occur in email, calendar, or task content — unlike the old
# "|||" / "===" which could appear in message bodies, subjects, or signatures
# and silently corrupt field parsing. They survive an `osascript -e` round-trip
# unchanged, and str.strip() only trims them at the very ends of the output
# (harmless, since empty records are filtered out during parsing).
DELIM = "\x1f"          # ASCII 31, Unit Separator — between fields
RECORD_DELIM = "\x1e"   # ASCII 30, Record Separator — between records
