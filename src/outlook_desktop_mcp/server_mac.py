"""
Outlook Desktop MCP Server — macOS
====================================
Exposes Microsoft Outlook for Mac as an MCP server over stdio.
Uses AppleScript automation via osascript — no Microsoft Graph, no Entra app.
Just run this on macOS with Outlook open and you have a full email MCP server.

Entry point: python -m outlook_desktop_mcp (auto-detected on macOS)
"""
import sys
import json
import logging
import os
import re

from mcp.server.fastmcp import FastMCP

from outlook_desktop_mcp.applescript_bridge import AppleScriptBridge
from outlook_desktop_mcp.utils.applescript_helpers import (
    escape,
    text_to_html,
    format_date,
    date_var_lines,
    parse_date,
    resolve_folder_ref,
    DELIM,
    RECORD_DELIM,
)
from datetime import datetime, timedelta

# --- Logging (all to stderr, stdout is reserved for MCP JSON-RPC) ---

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("outlook_desktop_mcp")

# --- MCP Server ---

mcp = FastMCP(
    "outlook-desktop-mcp",
    instructions=(
        "This MCP server gives you full access to Microsoft Outlook on macOS "
        "via AppleScript automation. It can send emails, read inbox messages, "
        "search across folders, mark messages as read/unread, move messages "
        "between folders, reply to emails, manage calendar events, and "
        "manage tasks.\n\n"
        "All operations use the locally running Outlook app — no "
        "Microsoft Graph API, no Entra app registration, no OAuth tokens needed. "
        "The user's existing Outlook session handles all authentication.\n\n"
        "PREREQUISITE: Microsoft Outlook for Mac must be running.\n\n"
        "NOTE: entry_id values on macOS are numeric IDs (not hex strings like "
        "on Windows). They identify items within their folder context.\n\n"
        "AVAILABLE TOOL CATEGORIES:\n"
        "- Email: send, draft, list, read, search, reply, forward, mark "
        "read/unread, move, snooze/unsnooze (follow-up flag + reminder)\n"
        "- Calendar: list calendars, list events, create appointments/meetings, "
        "update, delete, search events, respond to meeting invites\n"
        "  A profile usually holds many calendars across several accounts, and "
        "their names are not unique. Call list_calendars first and address a "
        "calendar by its numeric calendar_id, never by name.\n"
        "- Tasks: create, list, update, search, complete, delete to-do items\n"
        "- Categories: list color categories and set them on any item\n"
        "- Attachments: list and save attachments\n"
        "- Folders: list folder hierarchy"
    ),
)

bridge = AppleScriptBridge()


# --- Helper: truncate long text ---

def _truncate(text: str, max_length: int = 5000) -> str:
    if len(text) <= max_length:
        return text
    return text[:max_length] + "\n... [truncated]"


def _clean(value: str) -> str:
    """Replace AppleScript's 'missing value' with empty string."""
    v = value.strip()
    return "" if v == "missing value" else v


def _parse_due_bound(value: str, end_of_day: bool = False) -> datetime:
    """Parse an ISO 8601 date/datetime the caller passed as a due-date bound.

    A date with no time component is normalized to the start of that day, or —
    when end_of_day is set (for an inclusive upper bound) — to 23:59:59.
    """
    dt = datetime.fromisoformat(value)
    date_only = "T" not in value and ":" not in value
    if end_of_day and date_only:
        dt = dt.replace(hour=23, minute=59, second=59)
    return dt


def _due_to_datetime(due_raw: str) -> datetime | None:
    """Convert an AppleScript due-date string to a datetime, or None if it is
    empty or cannot be parsed."""
    if not due_raw:
        return None
    try:
        return datetime.fromisoformat(parse_date(due_raw))
    except ValueError:
        return None


def _recipient_lines(addresses: str, kind: str) -> str:
    """Build AppleScript `make new <kind> ...` lines from a semicolon-separated
    list of email addresses, attaching each to the message named `newMsg`."""
    lines = ""
    for addr in addresses.split(";"):
        addr = addr.strip()
        if addr:
            lines += (
                f'make new {kind} at newMsg with properties '
                f'{{email address:{{address:"{escape(addr)}"}}}}\n'
            )
    return lines


def _attachment_lines(paths: list[str] | None, target: str) -> str:
    """Build AppleScript lines that attach local files to `target`.

    Each path must be an existing absolute file. Raises FileNotFoundError with
    a descriptive message if any path is missing, so the caller can surface a
    clean error before running the script.
    """
    if not paths:
        return ""
    lines = ""
    for path in paths:
        path = path.strip()
        if not path:
            continue
        if not os.path.isfile(path):
            raise FileNotFoundError(f"Attachment not found: {path}")
        lines += (
            f'make new attachment at {target} with properties '
            f'{{file:POSIX file "{escape(path)}"}}\n'
        )
    return lines


# --- UI Scraping for New Outlook for Mac ---
# New Outlook for Mac stores Exchange/M365 mailbox data in the cloud and
# does NOT expose it through the AppleScript `inbox` keyword (which only
# reaches the empty local "On My Computer" inbox). The only way to access
# Exchange messages is via macOS UI scripting (System Events), reading
# the message list table visible in the Outlook window.


_UI_MESSAGE_LIST_PATH = (
    'tell application "System Events"\n'
    '    tell process "Microsoft Outlook"\n'
    '        tell window 1\n'
    '            tell splitter group 1\n'
    '                tell splitter group 1\n'
    '                    tell splitter group 1\n'
    '                        tell group 1\n'
    '                            tell scroll area 1\n'
    '                                tell table 1\n'
)

_UI_MESSAGE_LIST_END = (
    '                                end tell\n'
    '                            end tell\n'
    '                        end tell\n'
    '                    end tell\n'
    '                end tell\n'
    '            end tell\n'
    '        end tell\n'
    '    end tell\n'
    'end tell'
)


async def _ui_list_messages(bridge_obj, count: int = 10) -> list[dict]:
    """Read visible inbox messages via UI scripting (System Events).

    This is the fallback for New Outlook for Mac where AppleScript's inbox
    keyword only sees the empty local mailbox.
    """
    script = (
        _UI_MESSAGE_LIST_PATH +
        f'                                    set rowList to rows\n'
        f'                                    set rowCount to count of rowList\n'
        f'                                    set maxRows to rowCount\n'
        f'                                    if maxRows > {count} then set maxRows to {count}\n'
        f'                                    set output to ""\n'
        f'                                    repeat with i from 1 to maxRows\n'
        f'                                        set r to row i\n'
        f'                                        try\n'
        f'                                            set cellDesc to description of UI element 1 of r\n'
        f'                                            set output to output & cellDesc & "{RECORD_DELIM}"\n'
        f'                                        end try\n'
        f'                                    end repeat\n'
        f'                                    return output\n' +
        _UI_MESSAGE_LIST_END
    )

    raw = await bridge_obj.run(script)
    if not raw:
        return []

    results = []
    for idx, record in enumerate(raw.split(RECORD_DELIM), start=1):
        record = record.strip()
        if not record:
            continue
        # Cell description format uses `,` + 4+ spaces as major field
        # separators, while in-content commas have 0-1 trailing spaces.
        # Structure: [UNREAD_FLAG,]    SENDER, SUBJECT,     TIME,    [FLAGS,]
        fields = [f.strip() for f in re.split(r",\s{4,}", record)]

        is_unread = False
        has_attachment = False
        # Status tokens are locale-dependent. Match known tokens across
        # languages so the parser works regardless of macOS language.
        _UNREAD_TOKENS = {"Ulest", "Unread", "Non lu", "Nicht gelesen",
                          "No leído", "未読", "未读"}
        _ATTACHMENT_TOKENS = {"Har filer", "Has attachments", "Contient des fichiers",
                              "Hat Anlagen", "Tiene archivos adjuntos", "添付ファイルあり",
                              "有附件"}
        _SKIP_PREFIXES = ("Merket som", "Marked as", "Marqué comme",
                          "Markiert als", "Marcado como", "A ")
        _CATEGORY_TOKENS = {"Kategorisert", "Categorized", "Catégorisé",
                            "Kategorisiert", "Categorizado"}
        cleaned = []
        for f in fields:
            if not f:
                continue
            if f in _UNREAD_TOKENS:
                is_unread = True
                continue
            if any(tok in f for tok in _ATTACHMENT_TOKENS):
                has_attachment = True
                continue
            if f in _CATEGORY_TOKENS or any(f.startswith(p) for p in _SKIP_PREFIXES):
                continue
            cleaned.append(f)

        # cleaned is typically: [SENDER_AND_SUBJECT, TIME]
        # or [SENDER_AND_SUBJECT, TIME, extra...]
        # SENDER_AND_SUBJECT is: "Sender, Subject" (comma + 1 space)
        sender_subject = cleaned[0] if cleaned else ""
        time_str = cleaned[1] if len(cleaned) > 1 else ""
        # Strip trailing comma from time
        time_str = time_str.rstrip(",").strip()

        # Split sender from subject on first ", " (comma + single space)
        # Remove thread/unread count prefixes like "2 messages, " or
        # "1 unread message, " in any locale (pattern: digits + words + comma)
        ss = sender_subject
        ss = re.sub(r"^\d+\s+[\w\s]+,\s*", "", ss)
        # Split on first ", " to get sender and subject
        comma_pos = ss.find(", ")
        if comma_pos > 0:
            sender = ss[:comma_pos].strip()
            subject = ss[comma_pos + 2:].strip()
        else:
            sender = ""
            subject = ss.strip()

        results.append({
            "entry_id": f"ui-{idx}",
            "subject": subject or "(could not parse subject)",
            "sender": "",
            "sender_name": sender,
            "received_time": time_str,
            "unread": is_unread,
            "has_attachments": has_attachment,
            "attachment_count": 1 if has_attachment else 0,
            "_source": "ui_scraping",
        })

    return results


# =====================================================================
# TOOL 1: send_email
# =====================================================================

@mcp.tool()
async def send_email(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
    bcc: str = "",
    html_body: str = "",
    attachments: list[str] | None = None,
) -> str:
    """Send an email using the user's Outlook account.

    Creates and sends an email immediately through the default Outlook profile.
    The email will appear in the user's Sent Items folder after sending.

    Args:
        to: One or more recipient email addresses, separated by semicolons.
            Example: "alice@example.com" or "alice@example.com; bob@example.com"
        subject: The email subject line.
        body: The plain-text body of the email. Line breaks are preserved.
            Ignored when html_body is provided.
        cc: Optional. CC recipients, separated by semicolons.
        bcc: Optional. BCC recipients, separated by semicolons.
        html_body: Optional. HTML-formatted body, used verbatim instead of
            body. Supply a full fragment, e.g. "<p>Hello</p>".
        attachments: Optional. A list of absolute file paths to attach.

    Returns:
        A confirmation message with subject and recipients, or an error.
    """
    to_lines = _recipient_lines(to, "to recipient")
    cc_lines = _recipient_lines(cc, "cc recipient") if cc else ""
    bcc_lines = _recipient_lines(bcc, "bcc recipient") if bcc else ""

    try:
        att_lines = _attachment_lines(attachments, "newMsg")
    except FileNotFoundError as e:
        return json.dumps({"error": str(e)})

    # `content` is Outlook's HTML body property, so a plain-text body has to be
    # converted to HTML or its line breaks collapse. (There is no `html content`
    # property — naming it is an AppleScript syntax error.)
    content_prop = f'content:"{escape(html_body or text_to_html(body))}"'

    script = f'''tell application "Microsoft Outlook"
    set newMsg to make new outgoing message with properties {{subject:"{escape(subject)}", {content_prop}}}
    {to_lines}{cc_lines}{bcc_lines}{att_lines}
    send newMsg
end tell'''

    try:
        await bridge.run(script)
        return f"Email sent: '{subject}' to {to}"
    except Exception as e:
        return f"Error sending email: {e}"


# =====================================================================
# TOOL: create_draft
# =====================================================================

@mcp.tool()
async def create_draft(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
    bcc: str = "",
    html_body: str = "",
    attachments: list[str] | None = None,
) -> str:
    """Create an email draft without sending it.

    Builds the message and saves it to the Drafts folder so it can be
    reviewed and sent manually later. Useful when a human should approve
    the message before it goes out.

    Args:
        to: One or more recipient email addresses, separated by semicolons.
        subject: The email subject line.
        body: The plain-text body of the email. Line breaks are preserved.
        cc: Optional. CC recipients, separated by semicolons.
        bcc: Optional. BCC recipients, separated by semicolons.
        html_body: Optional. HTML-formatted body, used verbatim instead of body.
        attachments: Optional. A list of absolute file paths to attach.

    Returns:
        JSON with the draft's entry_id and subject, or an error.
    """
    to_lines = _recipient_lines(to, "to recipient")
    cc_lines = _recipient_lines(cc, "cc recipient") if cc else ""
    bcc_lines = _recipient_lines(bcc, "bcc recipient") if bcc else ""

    try:
        att_lines = _attachment_lines(attachments, "newMsg")
    except FileNotFoundError as e:
        return json.dumps({"error": str(e)})

    # `content` is Outlook's HTML body property, so a plain-text body has to be
    # converted to HTML or its line breaks collapse. (There is no `html content`
    # property — naming it is an AppleScript syntax error.)
    content_prop = f'content:"{escape(html_body or text_to_html(body))}"'

    # No explicit `save`: Outlook's AppleScript `save` verb demands an
    # `in <file>` parameter for outgoing messages and fails with -1701.
    # `make new outgoing message` already persists the item to Drafts.
    script = f'''tell application "Microsoft Outlook"
    set newMsg to make new outgoing message with properties {{subject:"{escape(subject)}", {content_prop}}}
    {to_lines}{cc_lines}{bcc_lines}{att_lines}
    return (id of newMsg as text) & "{DELIM}" & (subject of newMsg)
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM)
        result = {
            "status": "draft_created",
            "entry_id": parts[0].strip() if len(parts) > 0 else "",
            "subject": parts[1].strip() if len(parts) > 1 else subject,
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error creating draft: {e}"


# =====================================================================
# TOOL 2: list_emails
# =====================================================================

@mcp.tool()
async def list_emails(
    folder: str = "inbox",
    count: int = 10,
    unread_only: bool = False,
) -> str:
    """List recent emails from a specified Outlook folder.

    Returns a JSON array of email summaries sorted by received time (newest
    first). Each summary includes entry_id, subject, sender, sender_name,
    received_time, unread status, and attachment info.

    Use the entry_id from results to read full content with read_email,
    or to perform actions like mark_as_read, move_email, or reply_email.

    Args:
        folder: The folder to list. Case-insensitive names: "inbox" (default),
            "sent"/"sentmail", "drafts", "deleted"/"trash", "junk"/"spam",
            "outbox", or any custom folder name visible in list_folders output.
        count: Maximum number of emails to return. Default 10, max recommended 50.
        unread_only: If true, only return unread emails. Default false.

    Returns:
        JSON array of email summary objects.
    """
    folder_ref = resolve_folder_ref(folder)
    unread_filter = ' whose is read is false' if unread_only else ''

    script = f'''tell application "Microsoft Outlook"
    set folderRef to {folder_ref}
    set allMsgs to messages of folderRef{unread_filter}
    set msgCount to count of allMsgs
    set maxCount to {count}
    if msgCount < maxCount then set maxCount to msgCount
    set output to ""
    repeat with i from 1 to maxCount
        set m to item i of allMsgs
        set mid to id of m
        set msubject to subject of m
        set msender to ""
        try
            set msender to address of sender of m
        end try
        set msenderName to ""
        try
            set msenderName to name of sender of m
        end try
        set mtime to time received of m as string
        set misread to is read of m
        set mattcount to 0
        try
            set mattcount to count of attachments of m
        end try
        set output to output & (mid as text) & "{DELIM}" & msubject & "{DELIM}" & msender & "{DELIM}" & msenderName & "{DELIM}" & mtime & "{DELIM}" & (misread as text) & "{DELIM}" & (mattcount as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script)

        results = []
        if raw:
            for record in raw.split(RECORD_DELIM):
                record = record.strip()
                if not record:
                    continue
                parts = record.split(DELIM)
                if len(parts) < 7:
                    continue
                att_count = int(parts[6]) if parts[6].strip().isdigit() else 0
                results.append({
                    "entry_id": parts[0].strip(),
                    "subject": parts[1].strip() or "(no subject)",
                    "sender": parts[2].strip(),
                    "sender_name": parts[3].strip(),
                    "received_time": _clean(parts[4]),
                    "unread": parts[5].strip().lower() != "true",  # is_read -> unread
                    "has_attachments": att_count > 0,
                    "attachment_count": att_count,
                })

        # Fallback: New Outlook for Mac keeps Exchange messages outside the
        # AppleScript-visible mailbox. If the standard query returned nothing
        # for the inbox, try reading the visible message list via UI scripting.
        if not results and folder.lower().strip() == "inbox":
            try:
                results = await _ui_list_messages(bridge, count)
            except Exception:
                pass  # UI scraping failed — return empty list

        return json.dumps(results, indent=2, default=str)
    except Exception as e:
        return f"Error listing emails: {e}"


# =====================================================================
# TOOL 3: read_email
# =====================================================================

@mcp.tool()
async def read_email(
    entry_id: str = "",
    subject_search: str = "",
    folder: str = "inbox",
) -> str:
    """Read the full content of a specific email.

    Retrieves complete email details including body text, recipients, CC,
    and metadata. Provide EITHER entry_id (preferred, exact match) OR
    subject_search (finds most recent match by subject substring).

    Args:
        entry_id: The numeric ID of the email. Most reliable way to identify
            a specific email. Get this from list_emails or search_emails results.
        subject_search: Alternative to entry_id. A case-insensitive substring
            to search for in email subjects. Returns the most recent match.
        folder: Folder to search when using subject_search. Ignored when
            entry_id is provided. Default "inbox".

    Returns:
        JSON object with full email details (entry_id, subject, sender,
        sender_name, received_time, unread, to, cc, body, attachment info).
    """
    if entry_id:
        folder_ref = resolve_folder_ref(folder)
        script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set mid to id of m
    set msubject to subject of m
    set msender to ""
    try
        set msender to address of sender of m
    end try
    set msenderName to ""
    try
        set msenderName to name of sender of m
    end try
    set mtime to time received of m as string
    set misread to is read of m
    set mattcount to 0
    try
        set mattcount to count of attachments of m
    end try
    set mto to ""
    try
        set recips to to recipients of m
        repeat with r in recips
            set mto to mto & address of r & "; "
        end repeat
    end try
    set mcc to ""
    try
        set recips to cc recipients of m
        repeat with r in recips
            set mcc to mcc & address of r & "; "
        end repeat
    end try
    set mbody to ""
    try
        set mbody to plain text content of m
    end try
    return (mid as text) & "{DELIM}" & msubject & "{DELIM}" & msender & "{DELIM}" & msenderName & "{DELIM}" & mtime & "{DELIM}" & (misread as text) & "{DELIM}" & (mattcount as text) & "{DELIM}" & mto & "{DELIM}" & mcc & "{DELIM}" & mbody
end tell'''
    elif subject_search:
        folder_ref = resolve_folder_ref(folder)
        safe_query = escape(subject_search)
        script = f'''tell application "Microsoft Outlook"
    set folderRef to {folder_ref}
    set matchMsgs to messages of folderRef whose subject contains "{safe_query}"
    if (count of matchMsgs) = 0 then return "NOT_FOUND"
    set m to item 1 of matchMsgs
    set mid to id of m
    set msubject to subject of m
    set msender to ""
    try
        set msender to address of sender of m
    end try
    set msenderName to ""
    try
        set msenderName to name of sender of m
    end try
    set mtime to time received of m as string
    set misread to is read of m
    set mattcount to 0
    try
        set mattcount to count of attachments of m
    end try
    set mto to ""
    try
        set recips to to recipients of m
        repeat with r in recips
            set mto to mto & address of r & "; "
        end repeat
    end try
    set mcc to ""
    try
        set recips to cc recipients of m
        repeat with r in recips
            set mcc to mcc & address of r & "; "
        end repeat
    end try
    set mbody to ""
    try
        set mbody to plain text content of m
    end try
    return (mid as text) & "{DELIM}" & msubject & "{DELIM}" & msender & "{DELIM}" & msenderName & "{DELIM}" & mtime & "{DELIM}" & (misread as text) & "{DELIM}" & (mattcount as text) & "{DELIM}" & mto & "{DELIM}" & mcc & "{DELIM}" & mbody
end tell'''
    else:
        return json.dumps({"error": "Provide either entry_id or subject_search"})

    try:
        raw = await bridge.run(script)
        if raw == "NOT_FOUND":
            return json.dumps({"error": f"No email found matching '{subject_search}'"})

        parts = raw.split(DELIM, 9)  # max 10 parts
        if len(parts) < 10:
            return json.dumps({"error": "Failed to parse email data"})

        att_count = int(parts[6].strip()) if parts[6].strip().isdigit() else 0
        result = {
            "entry_id": parts[0].strip(),
            "subject": parts[1].strip() or "(no subject)",
            "sender": parts[2].strip(),
            "sender_name": parts[3].strip(),
            "received_time": _clean(parts[4]),
            "unread": parts[5].strip().lower() != "true",
            "has_attachments": att_count > 0,
            "attachment_count": att_count,
            "to": parts[7].strip(),
            "cc": parts[8].strip(),
            "body": _truncate(_clean(parts[9])),
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error reading email: {e}"


# =====================================================================
# TOOL 4: mark_as_read
# =====================================================================

@mcp.tool()
async def mark_as_read(entry_id: str) -> str:
    """Mark a specific email as read in Outlook.

    Changes the unread status to read, same as clicking on an email in Outlook.
    The change is persisted immediately and synced to the server.

    Args:
        entry_id: The numeric ID of the email. Get this from list_emails
            or search_emails results.

    Returns:
        Confirmation message with the email subject, or an error.
    """
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set is read of m to true
    return subject of m
end tell'''

    try:
        subject = await bridge.run(script)
        return f"Marked as read: '{subject}'"
    except Exception as e:
        return f"Error marking email as read: {e}"


# =====================================================================
# TOOL 5: mark_as_unread
# =====================================================================

@mcp.tool()
async def mark_as_unread(entry_id: str) -> str:
    """Mark a specific email as unread in Outlook.

    Restores a previously read email to unread status. Useful for flagging
    emails that need follow-up attention. Persisted immediately.

    Args:
        entry_id: The numeric ID of the email. Get this from list_emails
            or search_emails results.

    Returns:
        Confirmation message with the email subject, or an error.
    """
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set is read of m to false
    return subject of m
end tell'''

    try:
        subject = await bridge.run(script)
        return f"Marked as unread: '{subject}'"
    except Exception as e:
        return f"Error marking email as unread: {e}"


# =====================================================================
# TOOL 6: move_email
# =====================================================================

@mcp.tool()
async def move_email(
    entry_id: str,
    target_folder: str = "archive",
) -> str:
    """Move an email to a different Outlook folder.

    Moves the specified email from its current location to the target folder.
    IMPORTANT: After moving, the email gets a NEW entry_id — the old one
    becomes invalid.

    Args:
        entry_id: The numeric ID of the email to move.
        target_folder: Destination folder name. Default is "archive". Supports
            same names as list_emails: "inbox", "sent", "deleted"/"trash",
            "drafts", "junk"/"spam", or any custom folder name.

    Returns:
        Confirmation with email subject and destination, or an error.
    """
    dest_ref = resolve_folder_ref(target_folder)
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set msubject to subject of m
    move m to {dest_ref}
    return msubject
end tell'''

    try:
        subject = await bridge.run(script)
        return f"Moved '{subject}' to {target_folder}"
    except Exception as e:
        return f"Error moving email: {e}"


# =====================================================================
# TOOL: snooze_email
# =====================================================================

@mcp.tool()
async def snooze_email(
    entry_id: str,
    until: str,
    move_to_folder: str = "",
) -> str:
    """Postpone (snooze) an email: flag it for follow-up with a reminder.

    Sets the message's follow-up flag and schedules an Outlook reminder
    that pops up at the given time. Optionally also moves the message out
    of its current folder (e.g. to a "Snoozed"/"Pospuesto" folder) so the
    inbox stays clean until the reminder fires.

    NOTE: The message does NOT automatically return to the inbox at the
    reminder time — Outlook shows the reminder, and the flagged message
    also appears in the To-Do views. Use unsnooze_email to clear the flag
    (and optionally move it back to the inbox).

    Args:
        entry_id: The numeric ID of the email to snooze.
        until: When the reminder should fire, in ISO 8601 format.
            Examples: "2026-02-25 09:00", "2026-02-25T09:00:00".
        move_to_folder: Optional. A folder to move the message to while it
            is snoozed (same names as move_email, e.g. "Pospuesto").
            Top-level folders only. Leave empty to keep the message where
            it is. IMPORTANT: moving assigns a NEW entry_id, returned in
            the result.

    Returns:
        JSON with status, entry_id (the new one if moved), subject,
        reminder time, and folder, or an error.
    """
    try:
        until_dt = datetime.fromisoformat(until)
    except ValueError:
        return json.dumps({"error": f"Invalid ISO 8601 datetime: {until!r}"})

    date_lines = date_var_lines("remD", until_dt)
    move_line = ""
    if move_to_folder:
        move_line = f"set m to move m to {resolve_folder_ref(move_to_folder)}\n    "

    script = f'''tell application "Microsoft Outlook"
    {date_lines}
    set m to message id {entry_id}
    set todo flag of m to not completed
    set start date of m to remD
    set due date of m to remD
    set reminder date time of m to remD
    {move_line}return (id of m as text) & "{DELIM}" & (subject of m)
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM)
        result = {
            "status": "snoozed",
            "entry_id": parts[0].strip() if len(parts) > 0 else entry_id,
            "subject": parts[1].strip() if len(parts) > 1 else "",
            "reminder": until_dt.isoformat(),
            "folder": move_to_folder or "(unchanged)",
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error snoozing email: {e}"


# =====================================================================
# TOOL: unsnooze_email
# =====================================================================

@mcp.tool()
async def unsnooze_email(
    entry_id: str,
    move_to_inbox: bool = False,
) -> str:
    """Clear the snooze (follow-up flag and reminder) from an email.

    Removes the follow-up flag and cancels the pending reminder set by
    snooze_email. Optionally moves the message back to the inbox.

    Args:
        entry_id: The numeric ID of the email to unsnooze.
        move_to_inbox: If true, also move the message back to the inbox.
            IMPORTANT: moving assigns a NEW entry_id, returned in the result.

    Returns:
        JSON with status, entry_id (the new one if moved), and subject,
        or an error.
    """
    move_line = "set m to move m to inbox\n    " if move_to_inbox else ""

    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set todo flag of m to not flagged
    set start date of m to missing value
    set due date of m to missing value
    set reminder date time of m to missing value
    {move_line}return (id of m as text) & "{DELIM}" & (subject of m)
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM)
        result = {
            "status": "unsnoozed",
            "entry_id": parts[0].strip() if len(parts) > 0 else entry_id,
            "subject": parts[1].strip() if len(parts) > 1 else "",
            "folder": "inbox" if move_to_inbox else "(unchanged)",
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error unsnoozing email: {e}"


# =====================================================================
# TOOL 7: reply_email
# =====================================================================

@mcp.tool()
async def reply_email(
    entry_id: str,
    body: str,
    reply_all: bool = False,
    attachments: list[str] | None = None,
) -> str:
    """Reply to an email in Outlook.

    Creates and sends a reply, preserving the original message thread.
    Use reply_all=True to reply to all recipients (sender + CC list).

    Args:
        entry_id: The numeric ID of the email to reply to.
        body: The reply message text. Prepended above the original message
            in the email thread.
        reply_all: If true, reply to all recipients (sender + all CC/To).
            If false (default), reply only to the sender.
        attachments: Optional. A list of absolute file paths to attach.

    Returns:
        Confirmation indicating the reply was sent, or an error.
    """
    try:
        att_lines = _attachment_lines(attachments, "replyMsg")
    except FileNotFoundError as e:
        return json.dumps({"error": str(e)})

    reply_cmd = "reply all to" if reply_all else "reply to"
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set msubject to subject of m
    set replyMsg to {reply_cmd} m
    set content of replyMsg to "{escape(text_to_html(body, wrap=False))}<br><br>" & content of replyMsg
    {att_lines}send replyMsg
    return msubject
end tell'''

    try:
        subject = await bridge.run(script)
        return f"Reply sent to '{subject}' (reply_all={reply_all})"
    except Exception as e:
        return f"Error replying to email: {e}"


# =====================================================================
# TOOL: forward_email
# =====================================================================

@mcp.tool()
async def forward_email(
    entry_id: str,
    to: str,
    comment: str = "",
    cc: str = "",
    attachments: list[str] | None = None,
) -> str:
    """Forward an email to one or more recipients.

    Forwards the original message (including its existing attachments) and
    optionally prepends a comment above the forwarded content.

    Args:
        entry_id: The numeric ID of the email to forward.
        to: One or more recipient email addresses, separated by semicolons.
        comment: Optional. Text to prepend above the forwarded message.
        cc: Optional. CC recipients, separated by semicolons.
        attachments: Optional. Additional absolute file paths to attach.

    Returns:
        Confirmation indicating the message was forwarded, or an error.
    """
    to_lines = _recipient_lines(to, "to recipient")
    cc_lines = _recipient_lines(cc, "cc recipient") if cc else ""

    try:
        att_lines = _attachment_lines(attachments, "fwdMsg")
    except FileNotFoundError as e:
        return json.dumps({"error": str(e)})

    comment_line = ""
    if comment:
        comment_line = (
            f'set content of fwdMsg to '
            f'"{escape(text_to_html(comment, wrap=False))}<br><br>" '
            f'& content of fwdMsg\n'
        )

    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set msubject to subject of m
    set fwdMsg to forward m
    {to_lines}{cc_lines}{comment_line}{att_lines}send fwdMsg
    return msubject
end tell'''

    try:
        subject = await bridge.run(script)
        return f"Email forwarded: '{subject}' to {to}"
    except Exception as e:
        return f"Error forwarding email: {e}"


# =====================================================================
# TOOL 8: list_folders
# =====================================================================

@mcp.tool()
async def list_folders(max_depth: int = 2) -> str:
    """List all mail folders in the user's Outlook mailbox.

    Returns a JSON array showing the folder hierarchy with item counts.
    Use this to discover folder names for other tools (list_emails,
    move_email, search_emails).

    Args:
        max_depth: How many levels deep to recurse into subfolders.
            Default 2. Set to 1 for top-level only.

    Returns:
        JSON array of folder objects with name, item_count, and unread_count.
    """
    script = f'''tell application "Microsoft Outlook"
    set allFolders to mail folders
    set output to ""
    repeat with f in allFolders
        set fname to name of f
        set fcount to count of messages of f
        set funread to unread count of f
        set output to output & fname & "{DELIM}" & (fcount as text) & "{DELIM}" & (funread as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script)
        if not raw:
            return json.dumps([])

        results = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 3:
                continue
            results.append({
                "name": parts[0].strip(),
                "item_count": int(parts[1].strip()) if parts[1].strip().isdigit() else 0,
                "unread_count": int(parts[2].strip()) if parts[2].strip().isdigit() else 0,
            })
        return json.dumps(results, indent=2, default=str)
    except Exception as e:
        return f"Error listing folders: {e}"


# =====================================================================
# TOOL 9: search_emails
# =====================================================================

@mcp.tool()
async def search_emails(
    query: str,
    folder: str = "inbox",
    count: int = 10,
) -> str:
    """Search for emails in Outlook using text search.

    Searches email subjects using Outlook's AppleScript filtering.
    Results include entry_id for further operations.

    Args:
        query: The search term (case-insensitive substring match on subject).
            Examples: "budget report", "meeting notes", "quarterly".
        folder: Folder to search in. Default "inbox". Supports same
            names as list_emails.
        count: Maximum results to return. Default 10.

    Returns:
        JSON array of matching email summaries, or an error.
    """
    folder_ref = resolve_folder_ref(folder)
    safe_query = escape(query)

    script = f'''tell application "Microsoft Outlook"
    set folderRef to {folder_ref}
    set matchMsgs to messages of folderRef whose subject contains "{safe_query}"
    set msgCount to count of matchMsgs
    set maxCount to {count}
    if msgCount < maxCount then set maxCount to msgCount
    set output to ""
    repeat with i from 1 to maxCount
        set m to item i of matchMsgs
        set mid to id of m
        set msubject to subject of m
        set msender to ""
        try
            set msender to address of sender of m
        end try
        set msenderName to ""
        try
            set msenderName to name of sender of m
        end try
        set mtime to time received of m as string
        set misread to is read of m
        set mattcount to 0
        try
            set mattcount to count of attachments of m
        end try
        set output to output & (mid as text) & "{DELIM}" & msubject & "{DELIM}" & msender & "{DELIM}" & msenderName & "{DELIM}" & mtime & "{DELIM}" & (misread as text) & "{DELIM}" & (mattcount as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script)
        if not raw:
            return json.dumps([])

        results = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 7:
                continue
            att_count = int(parts[6].strip()) if parts[6].strip().isdigit() else 0
            results.append({
                "entry_id": parts[0].strip(),
                "subject": parts[1].strip() or "(no subject)",
                "sender": parts[2].strip(),
                "sender_name": parts[3].strip(),
                "received_time": _clean(parts[4]),
                "unread": parts[5].strip().lower() != "true",
                "has_attachments": att_count > 0,
                "attachment_count": att_count,
            })
        return json.dumps(results, indent=2, default=str)
    except Exception as e:
        return f"Error searching emails: {e}"


# =====================================================================
# CALENDAR TOOLS
# =====================================================================


# =====================================================================
# Calendar helpers
# =====================================================================

# AppleScript handlers that render a date as ISO 8601. Outlook returns dates
# coerced with `as string`, which follows the system locale ("miércoles, 23 de
# septiembre de 2026, 10:00:00" on a Spanish Mac) and is therefore neither
# parseable nor sortable. Building the string from numeric components sidesteps
# the locale entirely. These must sit at the top level of the script — outside
# any `tell` block — and are called as `my isoDate(...)`.
_ISO_PRELUDE = '''on pad2(n)
    set t to (n as integer) as text
    if (length of t) < 2 then set t to "0" & t
    return t
end pad2

on isoDate(d)
    if d is missing value then return ""
    return ((year of d) as text) & "-" & my pad2((month of d) as integer) & "-" & my pad2(day of d) & "T" & my pad2(hours of d) & ":" & my pad2(minutes of d) & ":" & my pad2(seconds of d)
end isoDate
'''

# Number of fields in one serialized event record, see _event_record_script.
_EVENT_FIELD_COUNT = 10


def _calendar_scope(calendar_id: str) -> str:
    """Return the AppleScript collection of events to query.

    An empty `calendar_id` means Outlook's global ``calendar events``
    collection, which spans every calendar of every account. A supplied id
    narrows it to that one calendar.

    Raises ValueError if the id is not a plain integer — the value is
    interpolated into the script, so it must never be free text.
    """
    cid = str(calendar_id).strip()
    if not cid:
        return "calendar events"
    return f"calendar events of calendar id {int(cid)}"


def _events_query(calendar_id: str, clauses: list[str]) -> str:
    """Build the event collection expression, optionally filtered.

    Outlook evaluates a `whose` clause internally, which is dramatically faster
    than fetching every event and filtering in Python: narrowing 2393 events to
    a 60-day window takes ~0.25s through `whose` versus ~44s iterating.
    """
    base = _calendar_scope(calendar_id)
    if not clauses:
        return base
    return f"({base} whose {' and '.join(clauses)})"


def _event_record_script(var: str = "e") -> str:
    """AppleScript appending one delimited record for event `var` to `output`.

    Field order matches _parse_event_record: id, subject, start, end, location,
    organizer, all-day flag, calendar id, calendar name, recurring flag.

    Every optional property is read inside its own `try` so that a single
    unreadable field degrades to an empty string instead of aborting the whole
    listing — Outlook exposes these inconsistently across account types.
    """
    return f'''        set esubj to ""
        try
            set esubj to (subject of {var}) as text
        end try
        set eloc to ""
        try
            set eloc to (location of {var}) as text
        end try
        set eorg to ""
        try
            set eorg to (organizer of {var}) as text
        end try
        set ecalId to ""
        set ecalName to ""
        try
            set ecalId to (id of (calendar of {var})) as text
            set ecalName to (name of (calendar of {var})) as text
        end try
        set erec to "false"
        try
            if (recurrence of {var}) is not missing value then set erec to "true"
        end try
        set output to output & ((id of {var}) as text) & "{DELIM}" & esubj & "{DELIM}" & my isoDate(start time of {var}) & "{DELIM}" & my isoDate(end time of {var}) & "{DELIM}" & eloc & "{DELIM}" & eorg & "{DELIM}" & ((all day flag of {var}) as text) & "{DELIM}" & ecalId & "{DELIM}" & ecalName & "{DELIM}" & erec & "{RECORD_DELIM}"
'''


def _parse_event_record(record: str) -> dict | None:
    """Parse one delimited event record, or None if it is malformed."""
    parts = record.split(DELIM)
    if len(parts) < _EVENT_FIELD_COUNT:
        return None
    return {
        "entry_id": parts[0].strip(),
        "subject": parts[1].strip() or "(no subject)",
        "start": parts[2].strip(),
        "end": parts[3].strip(),
        "location": _clean(parts[4]),
        "organizer": _clean(parts[5]),
        "all_day": parts[6].strip().lower() == "true",
        "calendar_id": parts[7].strip(),
        "calendar": _clean(parts[8]),
        "is_recurring": parts[9].strip().lower() == "true",
    }


def _parse_event_records(raw: str) -> list[dict]:
    """Parse a full AppleScript response into event dicts, sorted by start time.

    The ISO 8601 start strings sort lexicographically, so no date parsing is
    needed. Outlook returns events in no particular order.
    """
    events = []
    for record in raw.split(RECORD_DELIM):
        record = record.strip()
        if not record:
            continue
        parsed = _parse_event_record(record)
        if parsed is not None:
            events.append(parsed)
    events.sort(key=lambda ev: ev["start"])
    return events


# Calendar listings walk every event of every calendar, so they need more room
# than the bridge's 30s default.
_CALENDAR_TIMEOUT = 120


# =====================================================================
# TOOL: list_calendars
# =====================================================================

@mcp.tool()
async def list_calendars() -> str:
    """List every calendar Outlook knows about, across all accounts.

    Use this to discover calendar_id values for list_events, search_events and
    create_event. Calendar names are NOT unique — a profile can hold several
    calendars called "Calendar" in different accounts — so always target a
    calendar by its id, never by name.

    Outlook also exposes each account's root folder inside the calendar
    collection. Those are containers, not calendars: they have no name and hold
    no events. They are omitted here.

    Returns:
        JSON array of calendar objects with calendar_id, name, account and
        event_count. `account` is empty for calendars Outlook does not attribute
        to an account (shared and subscribed calendars, typically).
    """
    script = f'''tell application "Microsoft Outlook"
    set output to ""
    repeat with c in calendars
        set cname to missing value
        try
            set cname to name of c
        end try
        if cname is not missing value then
            set cacct to ""
            try
                set cacct to (name of (account of c)) as text
            end try
            set ccount to ""
            try
                set ccount to (count of (calendar events of c)) as text
            end try
            set output to output & ((id of c) as text) & "{DELIM}" & (cname as text) & "{DELIM}" & cacct & "{DELIM}" & ccount & "{RECORD_DELIM}"
        end if
    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script, timeout=_CALENDAR_TIMEOUT)
        if not raw:
            return json.dumps([])

        results = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 4:
                continue
            count_raw = parts[3].strip()
            results.append({
                "calendar_id": parts[0].strip(),
                "name": parts[1].strip(),
                "account": _clean(parts[2]),
                "event_count": int(count_raw) if count_raw.isdigit() else None,
            })
        return json.dumps(results, indent=2, default=str)
    except Exception as e:
        return f"Error listing calendars: {e}"


# =====================================================================
# TOOL 10: list_events
# =====================================================================

@mcp.tool()
async def list_events(
    start_date: str = "",
    end_date: str = "",
    count: int = 20,
    calendar_id: str = "",
) -> str:
    """List calendar events from Outlook within a date range.

    Returns a JSON array of event summaries sorted by start time. Each summary
    carries entry_id, subject, start, end, location, organizer, all_day, the
    calendar it lives in (calendar_id and calendar), and is_recurring.

    Use entry_id from results with get_event, update_event, or delete_event.

    By default this spans every calendar of every account. Pass calendar_id
    (from list_calendars) to list a single calendar.

    Note on recurring events: the date range is matched against each event's own
    start time, which for a recurring series is its FIRST occurrence. An ongoing
    weekly series that began before start_date is therefore not listed, even
    though it still has occurrences inside the range. Such events are flagged
    with is_recurring when they do appear.

    Args:
        start_date: Start of date range in ISO 8601 format (e.g. "2026-02-25"
            or "2026-02-25 09:00"). Default: now.
        end_date: End of date range. Default: 7 days from start_date.
        count: Maximum number of events to return, applied after sorting.
            Default 20.
        calendar_id: Optional. Restrict to one calendar, e.g. "133". Get ids
            from list_calendars. Default: all calendars.

    Returns:
        JSON array of event summary objects.
    """
    start = datetime.fromisoformat(start_date) if start_date else datetime.now()
    end = datetime.fromisoformat(end_date) if end_date else start + timedelta(days=7)

    try:
        scope = _events_query(calendar_id, [
            "start time is greater than or equal to startD",
            "start time is less than or equal to endD",
        ])
    except ValueError:
        return json.dumps({"error": f"Invalid calendar_id: {calendar_id!r}. Expected a number from list_calendars."})

    date_lines = date_var_lines("startD", start) + date_var_lines("endD", end)

    script = f'''{_ISO_PRELUDE}
tell application "Microsoft Outlook"
    {date_lines}
    set evts to {scope}
    set output to ""
    repeat with e in evts
{_event_record_script("e")}    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script, timeout=_CALENDAR_TIMEOUT)
        if not raw:
            return json.dumps([])
        return json.dumps(_parse_event_records(raw)[:count], indent=2, default=str)
    except Exception as e:
        return f"Error listing events: {e}"


# =====================================================================
# TOOL 11: get_event
# =====================================================================

@mcp.tool()
async def get_event(entry_id: str) -> str:
    """Read the full details of a specific calendar event.

    Retrieves complete event information including body/description,
    attendees, the calendar the event lives in, and whether it recurs.

    Args:
        entry_id: The numeric ID of the event. Get this from list_events
            or search_events results.

    Returns:
        JSON object with full event details.
    """
    script = f'''{_ISO_PRELUDE}
tell application "Microsoft Outlook"
    set e to calendar event id {entry_id}
    set eid to id of e
    set esubject to subject of e
    set elocation to ""
    try
        set elocation to location of e
    end try
    set eorganizer to ""
    try
        set eorganizer to organizer of e
    end try
    set eallday to all day flag of e
    set ebody to ""
    try
        set ebody to plain text content of e
    end try
    set eattendees to ""
    try
        set attList to attendees of e
        repeat with a in attList
            set eattendees to eattendees & address of a & "; "
        end repeat
    end try
    set ecalId to ""
    set ecalName to ""
    try
        set ecalId to (id of (calendar of e)) as text
        set ecalName to (name of (calendar of e)) as text
    end try
    set erec to "false"
    try
        if (recurrence of e) is not missing value then set erec to "true"
    end try
    return (eid as text) & "{DELIM}" & esubject & "{DELIM}" & my isoDate(start time of e) & "{DELIM}" & my isoDate(end time of e) & "{DELIM}" & elocation & "{DELIM}" & eorganizer & "{DELIM}" & (eallday as text) & "{DELIM}" & ebody & "{DELIM}" & eattendees & "{DELIM}" & ecalId & "{DELIM}" & ecalName & "{DELIM}" & erec
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM, 11)
        if len(parts) < 12:
            return json.dumps({"error": "Failed to parse event data"})

        result = {
            "entry_id": parts[0].strip(),
            "subject": parts[1].strip() or "(no subject)",
            "start": parts[2].strip(),
            "end": parts[3].strip(),
            "location": _clean(parts[4]),
            "organizer": _clean(parts[5]),
            "all_day": parts[6].strip().lower() == "true",
            "body": _truncate(_clean(parts[7])),
            "attendees": parts[8].strip(),
            "calendar_id": parts[9].strip(),
            "calendar": _clean(parts[10]),
            "is_recurring": parts[11].strip().lower() == "true",
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error reading event: {e}"


# =====================================================================
# TOOL 12: create_event
# =====================================================================

@mcp.tool()
async def create_event(
    subject: str,
    start: str,
    end: str,
    location: str = "",
    body: str = "",
    all_day: bool = False,
    reminder_minutes: int = 15,
    calendar_id: str = "",
) -> str:
    """Create a personal calendar appointment (no attendees).

    Creates and saves an appointment on the user's calendar. This is a
    personal event — no meeting invitations are sent. Use create_meeting
    instead if you need to invite attendees.

    Args:
        subject: The event title.
        start: Start time in ISO 8601 format. Examples: "2026-02-25 14:00",
            "2026-02-25T14:00:00". For all-day events, use just the date.
        end: End time in ISO 8601 format.
        location: Optional. Event location.
        body: Optional. Description or notes for the event.
        all_day: If true, creates an all-day event. Default false.
        reminder_minutes: Minutes before the event to show a reminder.
            Default 15. Set to 0 to disable reminder.
        calendar_id: Optional. Create the event in this calendar, e.g. "133".
            Get ids from list_calendars. Default: Outlook's default calendar.

    Returns:
        Confirmation with event subject, entry_id and the calendar it landed
        in, or an error.
    """
    start_dt = datetime.fromisoformat(start)
    end_dt = datetime.fromisoformat(end)

    try:
        cid = str(calendar_id).strip()
        target = f" at calendar id {int(cid)}" if cid else ""
    except ValueError:
        return json.dumps({"error": f"Invalid calendar_id: {calendar_id!r}. Expected a number from list_calendars."})

    date_lines = date_var_lines("startD", start_dt) + date_var_lines("endD", end_dt)
    props = f'subject:"{escape(subject)}", start time:startD, end time:endD'
    if location:
        props += f', location:"{escape(location)}"'
    if body:
        props += f', content:"{escape(text_to_html(body))}"'
    if all_day:
        props += ', all day flag:true'

    script = f'''{_ISO_PRELUDE}
tell application "Microsoft Outlook"
    {date_lines}
    set newEvt to make new calendar event{target} with properties {{{props}}}
    set ecalId to ""
    set ecalName to ""
    try
        set ecalId to (id of (calendar of newEvt)) as text
        set ecalName to (name of (calendar of newEvt)) as text
    end try
    return (id of newEvt as text) & "{DELIM}" & (subject of newEvt) & "{DELIM}" & my isoDate(start time of newEvt) & "{DELIM}" & my isoDate(end time of newEvt) & "{DELIM}" & ecalId & "{DELIM}" & ecalName
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM)
        result = {
            "status": "created",
            "entry_id": parts[0].strip() if len(parts) > 0 else "",
            "subject": parts[1].strip() if len(parts) > 1 else subject,
            "start": parts[2].strip() if len(parts) > 2 else start,
            "end": parts[3].strip() if len(parts) > 3 else end,
            "calendar_id": parts[4].strip() if len(parts) > 4 else "",
            "calendar": _clean(parts[5]) if len(parts) > 5 else "",
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error creating event: {e}"


# =====================================================================
# TOOL 13: create_meeting
# =====================================================================

@mcp.tool()
async def create_meeting(
    subject: str,
    start: str,
    end: str,
    required_attendees: str,
    location: str = "",
    body: str = "",
    optional_attendees: str = "",
    calendar_id: str = "",
) -> str:
    """Create a meeting and send invitations to attendees.

    Creates a calendar meeting and sends meeting requests to all specified
    attendees. Note: invitation delivery depends on Outlook's AppleScript
    meeting support on macOS.

    Args:
        subject: The meeting title.
        start: Start time in ISO 8601 format (e.g. "2026-02-25 14:00").
        end: End time in ISO 8601 format (e.g. "2026-02-25 15:00").
        required_attendees: Required attendee email addresses, separated by
            semicolons. Example: "alice@example.com; bob@example.com"
        location: Optional. Meeting location.
        body: Optional. Meeting description or agenda.
        optional_attendees: Optional. Optional attendee emails, separated
            by semicolons.
        calendar_id: Optional. Create the meeting in this calendar, e.g. "133".
            Get ids from list_calendars. Default: Outlook's default calendar.

    Returns:
        Confirmation that the meeting was created.
    """
    start_dt = datetime.fromisoformat(start)
    end_dt = datetime.fromisoformat(end)

    date_lines = date_var_lines("startD", start_dt) + date_var_lines("endD", end_dt)
    props = f'subject:"{escape(subject)}", start time:startD, end time:endD'
    if location:
        props += f', location:"{escape(location)}"'
    if body:
        props += f', content:"{escape(text_to_html(body))}"'

    try:
        cid = str(calendar_id).strip()
        target = f" at calendar id {int(cid)}" if cid else ""
    except ValueError:
        return json.dumps({"error": f"Invalid calendar_id: {calendar_id!r}. Expected a number from list_calendars."})

    attendee_lines = ""
    for addr in required_attendees.split(";"):
        addr = addr.strip()
        if addr:
            attendee_lines += f'make new required attendee at newEvt with properties {{email address:{{address:"{escape(addr)}"}}}}\n'
    if optional_attendees:
        for addr in optional_attendees.split(";"):
            addr = addr.strip()
            if addr:
                attendee_lines += f'make new optional attendee at newEvt with properties {{email address:{{address:"{escape(addr)}"}}}}\n'

    script = f'''tell application "Microsoft Outlook"
    {date_lines}
    set newEvt to make new calendar event{target} with properties {{{props}}}
    {attendee_lines}
    return (id of newEvt as text)
end tell'''

    try:
        await bridge.run(script)
        return f"Meeting '{subject}' created with attendees: {required_attendees}"
    except Exception as e:
        return f"Error creating meeting: {e}"


# =====================================================================
# TOOL 14: update_event
# =====================================================================

@mcp.tool()
async def update_event(
    entry_id: str,
    subject: str = "",
    start: str = "",
    end: str = "",
    location: str = "",
    body: str = "",
    all_day: bool | None = None,
) -> str:
    """Update an existing calendar event.

    Modifies properties of an appointment or meeting. Only the fields you
    provide will be updated — omitted fields remain unchanged.

    Args:
        entry_id: The numeric ID of the event to update.
        subject: Optional. New event title.
        start: Optional. New start time in ISO 8601 format.
        end: Optional. New end time in ISO 8601 format.
        location: Optional. New location.
        body: Optional. New description/notes.
        all_day: Optional. Pass false to turn an all-day event into a timed one
            (supply start and end too), or true to make it all-day. Omit to
            leave it as it is.

    Returns:
        Confirmation with updated event details, or an error.
    """
    set_lines = ""
    # The all-day flag is written before the times: clearing it on an all-day
    # event resets its start and end, which would otherwise discard the times
    # set in the same call.
    if all_day is not None:
        set_lines += f'set all day flag of e to {"true" if all_day else "false"}\n'
    if subject:
        set_lines += f'set subject of e to "{escape(subject)}"\n'
    if start:
        start_dt = datetime.fromisoformat(start)
        set_lines += date_var_lines("startD", start_dt)
        set_lines += 'set start time of e to startD\n'
    if end:
        end_dt = datetime.fromisoformat(end)
        set_lines += date_var_lines("endD", end_dt)
        set_lines += 'set end time of e to endD\n'
    if location:
        set_lines += f'set location of e to "{escape(location)}"\n'
    if body:
        set_lines += f'set content of e to "{escape(text_to_html(body))}"\n'

    if not set_lines:
        return json.dumps({"error": "No fields to update"})

    script = f'''{_ISO_PRELUDE}
tell application "Microsoft Outlook"
    set e to calendar event id {entry_id}
    {set_lines}
    return (id of e as text) & "{DELIM}" & (subject of e) & "{DELIM}" & my isoDate(start time of e) & "{DELIM}" & my isoDate(end time of e) & "{DELIM}" & (location of e) & "{DELIM}" & ((all day flag of e) as text)
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM)
        result = {
            "status": "updated",
            "entry_id": parts[0].strip() if len(parts) > 0 else entry_id,
            "subject": parts[1].strip() if len(parts) > 1 else "",
            "start": parts[2].strip() if len(parts) > 2 else "",
            "end": parts[3].strip() if len(parts) > 3 else "",
            "location": _clean(parts[4]) if len(parts) > 4 else "",
            "all_day": parts[5].strip().lower() == "true" if len(parts) > 5 else None,
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error updating event: {e}"


# =====================================================================
# TOOL 15: delete_event
# =====================================================================

@mcp.tool()
async def delete_event(entry_id: str) -> str:
    """Delete a calendar event.

    For personal appointments, the event is simply deleted. For meetings,
    the event is removed from your calendar.

    Args:
        entry_id: The numeric ID of the event to delete.

    Returns:
        Confirmation with the event subject, or an error.
    """
    script = f'''tell application "Microsoft Outlook"
    set e to calendar event id {entry_id}
    set esubject to subject of e
    delete e
    return esubject
end tell'''

    try:
        subject = await bridge.run(script)
        return f"Event deleted: '{subject}'"
    except Exception as e:
        return f"Error deleting event: {e}"


# =====================================================================
# TOOL: move_event
# =====================================================================

@mcp.tool()
async def move_event(entry_id: str, target_calendar_id: str) -> str:
    """Move a calendar event to a different calendar.

    REFUSES to move any event that has attendees. Outlook for Mac cannot
    reassign an event's calendar in place, so a move is a copy followed by
    deleting the original — and deleting a meeting sends a cancellation to
    everyone invited, then the copy invites them again. This tool will not do
    that to anyone. An event with attendees has to be moved by hand in Outlook,
    or left where it is.

    The copy preserves subject, times, location, body, organizer, categories
    and recurrence (verified against Outlook for Mac 16.113). The original is
    deleted only after the copy has been found and its subject verified, so a
    failed copy leaves the event untouched.

    Args:
        entry_id: The numeric ID of the event to move. Get it from
            list_events, search_events or get_event.
        target_calendar_id: The numeric ID of the destination calendar. Get it
            from list_calendars.

    Returns:
        JSON with the outcome: "moved" with the new entry_id, "refused" when
        the event has attendees, "unchanged", or "failed".
    """
    try:
        target = int(str(target_calendar_id).strip())
    except ValueError:
        return json.dumps({"error": f"Invalid target_calendar_id: {target_calendar_id!r}. Expected a number from list_calendars."})

    try:
        source_event = int(str(entry_id).strip())
    except ValueError:
        return json.dumps({"error": f"Invalid entry_id: {entry_id!r}. Expected a number."})

    # The copy is located by id, not by the command's result: Outlook's
    # `duplicate` returns nothing, so `set x to duplicate ...` leaves x
    # undefined. New events get increasing ids, so the copy is the highest id in
    # the target that was not there before — and its subject is checked against
    # the original before anything is deleted.
    #
    # The two lookups are narrowed with a `whose` clause (same subject, start
    # time within a minute) rather than walking the whole target calendar.
    # Walking it does not scale: moving into a 2300-event calendar blew past the
    # 120s timeout without completing a single move, while the narrowed query
    # answers in well under a second. Date equality inside `whose` silently
    # matches nothing in Outlook for Mac, hence the +/-60s window rather than
    # `start time is srcStart`.
    script = f"""tell application "Microsoft Outlook"
    set srcEv to calendar event id {source_event}
    set srcSubject to ""
    try
        set srcSubject to (subject of srcEv) as text
    end try
    set srcStart to missing value
    try
        set srcStart to start time of srcEv
    end try
    set srcEnd to missing value
    try
        set srcEnd to end time of srcEv
    end try
    set srcAllDay to false
    try
        set srcAllDay to all day flag of srcEv
    end try
    set srcLoc to missing value
    try
        set srcLoc to location of srcEv
    end try
    set srcContent to missing value
    try
        set srcContent to content of srcEv
    end try
    set srcCats to {{}}
    try
        set srcCats to category of srcEv
    end try
    set srcRecurring to false
    try
        if (recurrence of srcEv) is not missing value then set srcRecurring to true
    end try
    set srcCal to -1
    try
        set srcCal to id of (calendar of srcEv)
    end try

    set nAtt to 0
    try
        set nAtt to count of (attendees of srcEv)
    end try
    if nAtt > 0 then
        return "REFUSED{DELIM}" & (nAtt as text) & "{DELIM}" & srcSubject & "{DELIM}" & (srcCal as text)
    end if

    if srcCal is {target} then
        return "NOOP{DELIM}0{DELIM}" & srcSubject & "{DELIM}" & (srcCal as text)
    end if

    if srcStart is missing value then
        return "NOSTART{DELIM}0{DELIM}" & srcSubject & "{DELIM}" & (srcCal as text)
    end if
    set winLo to srcStart - 60
    set winHi to srcStart + 60

    set maxBefore to 0
    try
        repeat with e in (calendar events of calendar id {target} whose subject is srcSubject and start time is greater than or equal to winLo and start time is less than or equal to winHi)
            set thisId to id of e
            if thisId > maxBefore then set maxBefore to thisId
        end repeat
    end try

    duplicate srcEv to calendar id {target}

    set copyId to 0
    try
        repeat with e in (calendar events of calendar id {target} whose subject is srcSubject and start time is greater than or equal to winLo and start time is less than or equal to winHi)
            set thisId to id of e
            if thisId > maxBefore and thisId > copyId then set copyId to thisId
        end repeat
    end try

    set methodUsed to "duplicated"
    if copyId is 0 then
        -- Algunos calendarios (Gmail, compartidos) ignoran `duplicate` sin dar
        -- error: no se crea nada y no se avisa. Ahi se reconstruye el evento con
        -- `make new`, que si funciona. Se rechaza en los recurrentes porque
        -- reconstruir perderia la serie entera.
        if srcRecurring then
            return "NODUPRECUR{DELIM}0{DELIM}" & srcSubject & "{DELIM}" & (srcCal as text)
        end if
        if srcEnd is missing value then
            return "NOCOPY{DELIM}0{DELIM}" & srcSubject & "{DELIM}" & (srcCal as text)
        end if
        try
            if srcAllDay then
                set newEv to make new calendar event at calendar id {target} with properties {{subject:srcSubject, start time:srcStart, end time:srcEnd, all day flag:true}}
            else
                set newEv to make new calendar event at calendar id {target} with properties {{subject:srcSubject, start time:srcStart, end time:srcEnd}}
            end if
            try
                if srcLoc is not missing value then set location of newEv to srcLoc
            end try
            try
                if srcContent is not missing value then set content of newEv to srcContent
            end try
            try
                if (count of srcCats) > 0 then set category of newEv to srcCats
            end try
            set copyId to id of newEv
            set methodUsed to "recreated"
        on error errMsg
            return "NOCOPY{DELIM}0{DELIM}" & srcSubject & "{DELIM}" & (srcCal as text)
        end try
    end if
    if copyId is 0 then
        return "NOCOPY{DELIM}0{DELIM}" & srcSubject & "{DELIM}" & (srcCal as text)
    end if

    set copySubject to ""
    try
        set copySubject to (subject of (calendar event id copyId of calendar id {target})) as text
    end try
    if copySubject is not srcSubject then
        return "MISMATCH{DELIM}" & (copyId as text) & "{DELIM}" & copySubject & "{DELIM}" & (srcCal as text)
    end if

    delete (calendar event id {source_event} of calendar id srcCal)
    return "MOVED{DELIM}" & (copyId as text) & "{DELIM}" & srcSubject & "{DELIM}" & (srcCal as text) & "{DELIM}" & methodUsed
end tell"""

    try:
        raw = await bridge.run(script, timeout=_CALENDAR_TIMEOUT)
        parts = raw.split(DELIM)
        if len(parts) < 4:
            return json.dumps({"error": f"Unexpected response from Outlook: {raw!r}"})
        status = parts[0].strip()
        detail = parts[1].strip()
        subject = parts[2].strip()
        src_cal = parts[3].strip()

        if status == "REFUSED":
            return json.dumps({
                "status": "refused",
                "reason": "event has attendees",
                "attendees": int(detail) if detail.isdigit() else detail,
                "subject": subject,
                "entry_id": str(source_event),
                "calendar_id": src_cal,
                "detail": (
                    "Moving would delete the original, cancelling the meeting for "
                    "everyone invited, and then re-invite them from the copy. Move it "
                    "by hand in Outlook if that is acceptable."
                ),
            }, indent=2)
        if status == "NOOP":
            return json.dumps({
                "status": "unchanged",
                "reason": "already in that calendar",
                "subject": subject,
                "entry_id": str(source_event),
                "calendar_id": src_cal,
            }, indent=2)
        if status == "NODUPRECUR":
            return json.dumps({
                "status": "failed",
                "reason": ("the target calendar silently ignores `duplicate` (Gmail and shared "
                           "calendars do), and this event recurs, so rebuilding it would lose the "
                           "series; nothing was changed"),
                "subject": subject,
                "entry_id": str(source_event),
            }, indent=2)
        if status == "NOSTART":
            return json.dumps({
                "status": "failed",
                "reason": "the event has no readable start time, so the copy could not be located; nothing was changed",
                "subject": subject,
                "entry_id": str(source_event),
            }, indent=2)
        if status == "NOCOPY":
            return json.dumps({
                "status": "failed",
                "reason": "the copy never appeared; the original was left untouched",
                "subject": subject,
                "entry_id": str(source_event),
            }, indent=2)
        if status == "MISMATCH":
            return json.dumps({
                "status": "failed",
                "reason": ("the event found in the target did not match the original; the "
                           "original was left untouched, check the target calendar"),
                "subject": subject,
                "entry_id": str(source_event),
            }, indent=2)
        method = parts[4].strip() if len(parts) > 4 else "duplicated"
        result = {
            "status": "moved",
            "method": method,
            "subject": subject,
            "entry_id": detail,
            "previous_entry_id": str(source_event),
            "from_calendar_id": src_cal,
            "to_calendar_id": str(target),
            "note": "The event was copied and the original deleted, so its entry_id changed.",
        }
        if method == "recreated":
            result["note"] += (" The target calendar ignores `duplicate`, so the event was rebuilt "
                               "from its subject, times, all-day flag, location, body and categories.")
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Error moving event: {e}"


# =====================================================================
# TOOL 16: search_events
# =====================================================================

@mcp.tool()
async def search_events(
    query: str,
    start_date: str = "",
    end_date: str = "",
    count: int = 10,
    calendar_id: str = "",
) -> str:
    """Search for calendar events by keyword.

    Searches event subjects within a date range, across every calendar of every
    account unless calendar_id narrows it to one. Results are sorted by start
    time and carry the same fields as list_events, including the calendar each
    event lives in.

    The same recurring-event caveat as list_events applies: the date range is
    matched against each event's own start time, which for a series is its first
    occurrence.

    Args:
        query: The search term (case-insensitive substring match on subject).
            Examples: "standup", "review", "1:1".
        start_date: Start of search range in ISO 8601 format. Default: 30
            days ago.
        end_date: End of search range. Default: 30 days from now.
        count: Maximum results to return, applied after sorting. Default 10.
        calendar_id: Optional. Restrict to one calendar, e.g. "133". Get ids
            from list_calendars. Default: all calendars.

    Returns:
        JSON array of matching event summaries.
    """
    start = datetime.fromisoformat(start_date) if start_date else datetime.now() - timedelta(days=30)
    end = datetime.fromisoformat(end_date) if end_date else datetime.now() + timedelta(days=30)

    try:
        scope = _events_query(calendar_id, [
            f'subject contains "{escape(query)}"',
            "start time is greater than or equal to startD",
            "start time is less than or equal to endD",
        ])
    except ValueError:
        return json.dumps({"error": f"Invalid calendar_id: {calendar_id!r}. Expected a number from list_calendars."})

    date_lines = date_var_lines("startD", start) + date_var_lines("endD", end)

    script = f'''{_ISO_PRELUDE}
tell application "Microsoft Outlook"
    {date_lines}
    set evts to {scope}
    set output to ""
    repeat with e in evts
{_event_record_script("e")}    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script, timeout=_CALENDAR_TIMEOUT)
        if not raw:
            return json.dumps([])
        return json.dumps(_parse_event_records(raw)[:count], indent=2, default=str)
    except Exception as e:
        return f"Error searching events: {e}"


# =====================================================================
# TOOL: respond_to_meeting
# =====================================================================

_MEETING_RESPONSE_VERB = {
    "accept": "accept invite",
    "tentative": "accept tentatively invite",
    "decline": "decline invite",
}


@mcp.tool()
async def respond_to_meeting(
    entry_id: str,
    response: str,
    send_response: bool = True,
    comment: str = "",
) -> str:
    """Respond to a meeting invitation (accept, decline, or tentative).

    Operates on a meeting invite message in your mailbox. The meeting is
    added to (or updated on) your calendar accordingly, and — unless
    send_response is false — a response is sent to the organizer.

    Args:
        entry_id: The numeric ID of the meeting-invite message. Note this is
            the id of the invite message, not a calendar event id.
        response: Your response. One of: "accept", "decline", or "tentative".
        send_response: If true (default), send your response to the organizer.
            If false, update your calendar without notifying the organizer.
        comment: Optional. A note to include with the response (sent to the
            organizer). Ignored when send_response is false.

    Returns:
        Confirmation of your response, or an error.
    """
    resp = response.lower().strip()
    verb = _MEETING_RESPONSE_VERB.get(resp)
    if verb is None:
        return json.dumps({"error": f"Invalid response: {response!r}. Use accept, decline, or tentative."})

    send_flag = "true" if send_response else "false"
    comment_part = f' comment "{escape(comment)}"' if (comment and send_response) else ""

    script = f'''tell application "Microsoft Outlook"
    set mm to meeting message id {entry_id}
    set subj to subject of mm
    {verb} mm sending response {send_flag}{comment_part}
    return subj
end tell'''

    try:
        subj = await bridge.run(script)
        return f"Responded '{resp}' to meeting: '{subj}'"
    except Exception as e:
        return f"Error responding to meeting: {e}"


# =====================================================================
# TASK TOOLS
# =====================================================================

@mcp.tool()
async def list_tasks(
    include_completed: bool = False,
    count: int = 20,
    due_start: str = "",
    due_end: str = "",
) -> str:
    """List tasks from the Outlook Tasks folder.

    Returns a JSON array of task summaries. Each task includes entry_id,
    subject, due_date, and completion status.

    Args:
        include_completed: If true, include completed tasks. Default false
            (only pending/in-progress tasks).
        count: Maximum number of tasks to return. Default 20.
        due_start: Optional. Only return tasks due on or after this date
            (ISO 8601, e.g. "2026-03-01"). Tasks with no due date are excluded
            when any due filter is set.
        due_end: Optional. Only return tasks due on or before this date. A
            date with no time is treated as inclusive of that whole day.

    Returns:
        JSON array of task summary objects.
    """
    ds = _parse_due_bound(due_start) if due_start else None
    de = _parse_due_bound(due_end, end_of_day=True) if due_end else None
    filtering = ds is not None or de is not None

    completed_filter = "" if include_completed else " whose todo flag is not completed"
    # Overfetch when filtering, since the count cap is applied after date filtering.
    fetch_limit = count * 5 if filtering else count

    script = f'''tell application "Microsoft Outlook"
    set taskList to tasks{completed_filter}
    set taskCount to count of taskList
    set maxCount to {fetch_limit}
    if taskCount < maxCount then set maxCount to taskCount
    set output to ""
    repeat with i from 1 to maxCount
        set t to item i of taskList
        set tid to id of t
        set tname to name of t
        set tdue to ""
        try
            set tdue to due date of t as string
        end try
        set tflag to todo flag of t
        set tpriority to priority of t
        set output to output & (tid as text) & "{DELIM}" & tname & "{DELIM}" & tdue & "{DELIM}" & (tflag as text) & "{DELIM}" & (tpriority as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script)
        if not raw:
            return json.dumps([])

        results = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 5:
                continue
            due_raw = _clean(parts[2])
            if filtering:
                due_dt = _due_to_datetime(due_raw)
                if due_dt is None:
                    continue
                if ds is not None and due_dt < ds:
                    continue
                if de is not None and due_dt > de:
                    continue
            results.append({
                "entry_id": parts[0].strip(),
                "subject": parts[1].strip() or "(no subject)",
                "due_date": due_raw or None,
                "complete": parts[3].strip() == "completed",
                "priority": parts[4].strip(),
            })
            if len(results) >= count:
                break
        return json.dumps(results, indent=2, default=str)
    except Exception as e:
        return f"Error listing tasks: {e}"


@mcp.tool()
async def get_task(entry_id: str) -> str:
    """Read the full details of a specific task.

    Args:
        entry_id: The numeric ID of the task.

    Returns:
        JSON object with full task details including body.
    """
    script = f'''tell application "Microsoft Outlook"
    set t to task id {entry_id}
    set tid to id of t
    set tname to name of t
    set tdue to ""
    try
        set tdue to due date of t as string
    end try
    set tflag to todo flag of t
    set tpriority to priority of t
    set tbody to ""
    try
        set tbody to plain text content of t
    end try
    set tstartdate to ""
    try
        set tstartdate to start date of t as string
    end try
    return (tid as text) & "{DELIM}" & tname & "{DELIM}" & tdue & "{DELIM}" & (tflag as text) & "{DELIM}" & (tpriority as text) & "{DELIM}" & tbody & "{DELIM}" & tstartdate
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM, 6)
        if len(parts) < 7:
            return json.dumps({"error": "Failed to parse task data"})

        result = {
            "entry_id": parts[0].strip(),
            "subject": parts[1].strip() or "(no subject)",
            "due_date": _clean(parts[2]) or None,
            "complete": parts[3].strip() == "completed",
            "priority": parts[4].strip(),
            "body": _truncate(_clean(parts[5])),
            "start_date": _clean(parts[6]) or None,
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error reading task: {e}"


@mcp.tool()
async def create_task(
    subject: str,
    body: str = "",
    due_date: str = "",
    importance: str = "normal",
) -> str:
    """Create a new task in Outlook.

    Args:
        subject: The task title.
        body: Optional. Task description or notes.
        due_date: Optional. Due date in ISO 8601 format (e.g. "2026-03-01").
        importance: Optional. "low", "normal" (default), or "high".

    Returns:
        Confirmation with task subject and entry_id.
    """
    imp_map = {"low": "priority low", "normal": "priority normal", "high": "priority high"}
    imp_val = imp_map.get(importance.lower(), "priority normal")

    date_lines = ""
    props = f'name:"{escape(subject)}", priority:{imp_val}'
    if due_date:
        due_dt = datetime.fromisoformat(due_date)
        date_lines += date_var_lines("dueD", due_dt)
        props += ', due date:dueD'
    if body:
        props += f', content:"{escape(text_to_html(body))}"'

    script = f'''tell application "Microsoft Outlook"
    {date_lines}
    set newTask to make new task with properties {{{props}}}
    return (id of newTask as text) & "{DELIM}" & (name of newTask)
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM)
        result = {
            "status": "created",
            "entry_id": parts[0].strip() if len(parts) > 0 else "",
            "subject": parts[1].strip() if len(parts) > 1 else subject,
            "due_date": due_date or None,
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error creating task: {e}"


@mcp.tool()
async def update_task(
    entry_id: str,
    subject: str = "",
    body: str = "",
    due_date: str = "",
    start_date: str = "",
    importance: str = "",
    complete: bool | None = None,
) -> str:
    """Update an existing task.

    Modifies properties of a task. Only the fields you provide will be
    updated — omitted fields remain unchanged.

    Args:
        entry_id: The numeric ID of the task to update.
        subject: Optional. New task title.
        body: Optional. New description/notes.
        due_date: Optional. New due date in ISO 8601 format (e.g. "2026-03-01").
        start_date: Optional. New start date in ISO 8601 format.
        importance: Optional. "low", "normal", or "high".
        complete: Optional. Set true to mark complete, false to reopen.

    Returns:
        Confirmation with updated task details, or an error.
    """
    imp_map = {"low": "priority low", "normal": "priority normal", "high": "priority high"}

    set_lines = ""
    if subject:
        set_lines += f'set name of t to "{escape(subject)}"\n'
    if body:
        set_lines += f'set content of t to "{escape(text_to_html(body))}"\n'
    if due_date:
        due_dt = datetime.fromisoformat(due_date)
        set_lines += date_var_lines("dueD", due_dt)
        set_lines += 'set due date of t to dueD\n'
    if start_date:
        start_dt = datetime.fromisoformat(start_date)
        set_lines += date_var_lines("startD", start_dt)
        set_lines += 'set start date of t to startD\n'
    if importance:
        imp_val = imp_map.get(importance.lower())
        if imp_val is None:
            return json.dumps({"error": f"Invalid importance: {importance!r}. Use low, normal, or high."})
        set_lines += f'set priority of t to {imp_val}\n'
    if complete is not None:
        flag = "completed" if complete else "not completed"
        set_lines += f'set todo flag of t to {flag}\n'

    if not set_lines:
        return json.dumps({"error": "No fields to update"})

    script = f'''tell application "Microsoft Outlook"
    set t to task id {entry_id}
    {set_lines}
    return (id of t as text) & "{DELIM}" & (name of t) & "{DELIM}" & (todo flag of t as text) & "{DELIM}" & (priority of t as text)
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM)
        result = {
            "status": "updated",
            "entry_id": parts[0].strip() if len(parts) > 0 else entry_id,
            "subject": parts[1].strip() if len(parts) > 1 else "",
            "complete": parts[2].strip() == "completed" if len(parts) > 2 else None,
            "priority": parts[3].strip() if len(parts) > 3 else "",
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error updating task: {e}"


@mcp.tool()
async def search_tasks(
    query: str,
    include_completed: bool = False,
    count: int = 20,
) -> str:
    """Search tasks by keyword in their title.

    Args:
        query: The search term (case-insensitive substring match on the task
            title). Examples: "invoice", "follow up", "call".
        include_completed: If true, include completed tasks. Default false.
        count: Maximum number of tasks to return. Default 20.

    Returns:
        JSON array of matching task summary objects.
    """
    safe_query = escape(query)
    completed_filter = "" if include_completed else " and todo flag is not completed"

    script = f'''tell application "Microsoft Outlook"
    set taskList to tasks whose name contains "{safe_query}"{completed_filter}
    set taskCount to count of taskList
    set maxCount to {count}
    if taskCount < maxCount then set maxCount to taskCount
    set output to ""
    repeat with i from 1 to maxCount
        set t to item i of taskList
        set tid to id of t
        set tname to name of t
        set tdue to ""
        try
            set tdue to due date of t as string
        end try
        set tflag to todo flag of t
        set tpriority to priority of t
        set output to output & (tid as text) & "{DELIM}" & tname & "{DELIM}" & tdue & "{DELIM}" & (tflag as text) & "{DELIM}" & (tpriority as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script)
        if not raw:
            return json.dumps([])

        results = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 5:
                continue
            results.append({
                "entry_id": parts[0].strip(),
                "subject": parts[1].strip() or "(no subject)",
                "due_date": _clean(parts[2]) or None,
                "complete": parts[3].strip() == "completed",
                "priority": parts[4].strip(),
            })
        return json.dumps(results, indent=2, default=str)
    except Exception as e:
        return f"Error searching tasks: {e}"


@mcp.tool()
async def complete_task(entry_id: str) -> str:
    """Mark a task as complete.

    Sets the task status to complete.

    Args:
        entry_id: The numeric ID of the task.

    Returns:
        Confirmation with the task subject.
    """
    script = f'''tell application "Microsoft Outlook"
    set t to task id {entry_id}
    set todo flag of t to completed
    return name of t
end tell'''

    try:
        name = await bridge.run(script)
        return f"Task completed: '{name}'"
    except Exception as e:
        return f"Error completing task: {e}"


@mcp.tool()
async def delete_task(entry_id: str) -> str:
    """Delete a task from Outlook.

    Args:
        entry_id: The numeric ID of the task to delete.

    Returns:
        Confirmation with the task subject.
    """
    script = f'''tell application "Microsoft Outlook"
    set t to task id {entry_id}
    set tname to name of t
    delete t
    return tname
end tell'''

    try:
        name = await bridge.run(script)
        return f"Task deleted: '{name}'"
    except Exception as e:
        return f"Error deleting task: {e}"


# =====================================================================
# CATEGORY TOOLS
# =====================================================================

@mcp.tool()
async def list_categories() -> str:
    """List all available Outlook color categories.

    Returns the categories configured in the user's Outlook profile. These
    can be applied to emails, tasks, and events via set_category.

    Returns:
        JSON array of category objects with name and color.
    """
    script = f'''tell application "Microsoft Outlook"
    set output to ""
    repeat with c in categories
        set cname to name of c
        set ccolor to ""
        try
            set ccolor to color of c as text
        end try
        set output to output & cname & "{DELIM}" & ccolor & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script)
        if not raw:
            return json.dumps([])

        results = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            results.append({
                "name": parts[0].strip(),
                "color": parts[1].strip() if len(parts) > 1 else "",
            })
        return json.dumps(results, indent=2, default=str)
    except Exception as e:
        return f"Error listing categories: {e}"


_ITEM_REF = {
    "email": "message",
    "message": "message",
    "task": "task",
    "event": "calendar event",
    "calendar": "calendar event",
}


@mcp.tool()
async def set_category(entry_id: str, categories: str, item_type: str) -> str:
    """Set color categories on an email, task, or event.

    Replaces any existing categories on the item. A category that does not
    exist yet is created automatically (mirroring Outlook's behavior).

    IMPORTANT: On macOS the numeric entry_id is only unique within an item
    type — the same number can refer to a different email, task, and event —
    so item_type is required to target the correct item.

    Args:
        entry_id: The numeric ID of the item.
        categories: Category name(s), comma-separated. Example: "Important"
            or "Work, Follow-up". Use an empty string to clear all categories.
        item_type: The kind of item the entry_id refers to: "email", "task",
            or "event".

    Returns:
        Confirmation with the item name and applied categories, or an error.
    """
    ref = _ITEM_REF.get(item_type.lower().strip())
    if ref is None:
        return json.dumps({"error": f"Invalid item_type: {item_type!r}. Use email, task, or event."})

    names = [n.strip() for n in categories.split(",") if n.strip()]
    # Build an AppleScript list literal of the requested category names.
    list_literal = "{" + ", ".join(f'"{escape(n)}"' for n in names) + "}"

    script = f'''tell application "Microsoft Outlook"
    set theItem to {ref} id {entry_id}
    set catList to {{}}
    repeat with nm in {list_literal}
        set nmT to nm as text
        set matches to (categories whose name is nmT)
        if (count of matches) > 0 then
            set end of catList to item 1 of matches
        else
            set end of catList to (make new category with properties {{name:nmT}})
        end if
    end repeat
    set category of theItem to catList
    set itemName to ""
    try
        set itemName to subject of theItem
    end try
    if itemName is "" then
        try
            set itemName to name of theItem
        end try
    end if
    return itemName
end tell'''

    try:
        name = await bridge.run(script)
        applied = ", ".join(names) if names else "(none)"
        return f"Categories set on '{name}': {applied}"
    except Exception as e:
        return f"Error setting categories: {e}"


# =====================================================================
# ATTACHMENT TOOLS
# =====================================================================

@mcp.tool()
async def list_attachments(entry_id: str) -> str:
    """List all attachments on an email.

    Args:
        entry_id: The numeric ID of the email to check for attachments.

    Returns:
        JSON array of attachment objects with index and filename.
    """
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set attList to attachments of m
    set attCount to count of attList
    set output to ""
    repeat with i from 1 to attCount
        set a to item i of attList
        set aname to name of a
        set asize to file size of a
        set output to output & (i as text) & "{DELIM}" & aname & "{DELIM}" & (asize as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script)
        if not raw:
            return json.dumps([])

        results = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 3:
                continue
            results.append({
                "index": int(parts[0].strip()) if parts[0].strip().isdigit() else 0,
                "filename": parts[1].strip(),
                "size": int(parts[2].strip()) if parts[2].strip().isdigit() else 0,
            })
        return json.dumps(results, indent=2, default=str)
    except Exception as e:
        return f"Error listing attachments: {e}"


@mcp.tool()
async def save_attachment(
    entry_id: str,
    attachment_index: int = 1,
    save_directory: str = "",
) -> str:
    """Save an attachment from an email to disk.

    Downloads the specified attachment to a local directory.

    Args:
        entry_id: The numeric ID of the email containing the attachment.
        attachment_index: Which attachment to save (1-based index). Default 1.
            Use list_attachments to see available indices.
        save_directory: Directory to save the file to. Default: user's
            Downloads folder.

    Returns:
        The full file path where the attachment was saved, or an error.
    """
    if not save_directory:
        save_directory = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(save_directory, exist_ok=True)

    # Use POSIX path for AppleScript
    save_dir_posix = save_directory

    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set attList to attachments of m
    set attCount to count of attList
    if attCount < {attachment_index} then return "ERROR:Only " & attCount & " attachment(s), requested index {attachment_index}"
    set a to item {attachment_index} of attList
    set aname to name of a
    set savePath to POSIX file "{escape(save_dir_posix)}/{escape("__PLACEHOLDER__")}"
    save a in file ((POSIX path of (POSIX file "{escape(save_dir_posix)}")) & aname)
    return aname
end tell'''

    # Simpler approach: save to known path
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set attList to attachments of m
    set attCount to count of attList
    if attCount < {attachment_index} then return "ERROR:Only " & attCount & " attachment(s)"
    set a to item {attachment_index} of attList
    set aname to name of a
    set savePath to "{escape(save_dir_posix)}/" & aname
    save a in savePath
    return aname & "{DELIM}" & savePath
end tell'''

    try:
        raw = await bridge.run(script)
        if raw.startswith("ERROR:"):
            return raw

        parts = raw.split(DELIM)
        filename = parts[0].strip() if len(parts) > 0 else "unknown"
        save_path = os.path.join(save_directory, filename)
        result = {
            "status": "saved",
            "filename": filename,
            "path": save_path,
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error saving attachment: {e}"


# =====================================================================
# Entry point
# =====================================================================

def main():
    import asyncio

    async def _start():
        logger.info("Starting Outlook Desktop MCP server (macOS)...")
        await bridge.start()
        logger.info("AppleScript bridge ready. Starting MCP stdio transport...")

    asyncio.run(_start())
    try:
        mcp.run(transport="stdio")
    finally:
        bridge.stop()
