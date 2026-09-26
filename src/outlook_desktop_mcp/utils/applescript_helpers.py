"""Helpers for building and parsing AppleScript safely."""
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


def text_to_html(text: str, wrap: bool = True) -> str:
    """Convert plain text to HTML so its line breaks survive.

    Outlook for Mac's ``content`` property is documented as *the HTML content*
    of a message, event, or task — assigning raw plain text to it makes Outlook
    treat that text as markup, so newlines collapse into ordinary whitespace
    and the body renders as one run-on paragraph. The text is therefore
    HTML-escaped and its newlines turned into ``<br>`` before assignment.

    Args:
        text: The plain text to convert.
        wrap: When true (default) the result is wrapped in
            ``<html><body>…</body></html>``. Pass false when the fragment is
            being prepended to content that already carries those tags, such as
            the quoted original in a reply or forward.

    Returns an empty string for empty input, so callers can keep skipping the
    property entirely when there is no body.
    """
    if not text:
        return ""
    escaped = (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )
    # Normalise CRLF / CR to LF first so no line break is emitted twice.
    escaped = escaped.replace("\r\n", "\n").replace("\r", "\n")
    html = escaped.replace("\n", "<br>")
    return f"<html><body>{html}</body></html>" if wrap else html


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
# unchanged. Python's str.strip() counts both as whitespace, though, so output
# must never be stripped before it is split: an empty trailing field would
# vanish along with its delimiter.
DELIM = "\x1f"          # ASCII 31, Unit Separator — between fields
RECORD_DELIM = "\x1e"   # ASCII 30, Record Separator — between records
