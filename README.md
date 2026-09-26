# outlook-desktop-mcp

[![PyPI](https://img.shields.io/pypi/v/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Python](https://img.shields.io/pypi/pyversions/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-blue)]()

**Turn your running Outlook Desktop into an MCP server.** No Microsoft Graph API, no Entra app registration, no OAuth tokens — just your local Outlook and the authentication you already have.

Any MCP client (Claude Code, Claude Desktop, etc.) can then send emails, manage your calendar, create tasks, handle attachments, and more — all through your existing Outlook session.

## Quick Start

**1. Install** (requires Python 3.12+):

```bash
pip install outlook-desktop-mcp
```

**2. Register with Claude Code:**

```bash
claude mcp add outlook-desktop -- outlook-desktop-mcp
```

**3. Open Outlook and start a Claude Code session.** That's it — tools are available immediately.

## How It Works — Platform Routing

When the server starts, it checks which operating system it is running on and takes one of two paths:

```
                        outlook-desktop-mcp starts
                                  |
                          sys.platform check
                         /                  \
                   "win32"                "darwin"
                      |                      |
              ┌───────┴────────┐    ┌────────┴────────┐
              │  server.py     │    │  server_mac.py   │
              │  COM Bridge    │    │  AppleScript     │
              │  (34 tools)    │    │  Bridge          │
              │                │    │  (34 tools)      │
              └───────┬────────┘    └────────┬─────────┘
                      |                      |
              OUTLOOK.EXE via         Microsoft Outlook
              COM / STA thread        via osascript
                      |                      |
              Exchange / M365         Exchange / M365
```

**Both paths use your locally running Outlook app and its existing authenticated session.** No cloud credentials, no Graph API tokens — the server inherits whatever account Outlook is signed into.

### Why two paths?

Windows Outlook (Classic) exposes a rich COM automation interface — the Outlook Object Model (`MSOUTL.OLB`). This has been the standard way to programmatically control Outlook on Windows for over 20 years. It provides deep access to mail rules, categories, MAPI properties, and the full folder hierarchy.

Mac Outlook does not support COM. Instead, it exposes an AppleScript dictionary that can be driven via the `osascript` command. The AppleScript interface covers the core operations — email, calendar, tasks, categories — but does not expose mail rules, Out of Office, or certain advanced MAPI features. This is a limitation of what Microsoft chose to include in Outlook for Mac's scripting dictionary, not a limitation of this project.

The server is structured as two parallel implementations. 30 tools are shared under the same names, and each platform registers 4 more that only it can support (see the tables below). Shared tools mostly take the same parameters, but each side has a few extras the other lacks — for example `account` on Windows and `calendar_id` on macOS — so check a tool's schema rather than assuming the two are identical.

## Requirements

### Windows

- **Outlook Desktop (Classic)** — the `OUTLOOK.EXE` that comes with Microsoft 365 / Office. The new "modern" Outlook (`olk.exe`) does **not** support COM
- **Python 3.12+** (x64 or ARM64)
- **Outlook must be running** when the MCP server starts

Both x64 and ARM64 Windows are supported. On ARM64, all dependencies (`pywin32`, `mcp`, `pydantic-core`, `cryptography`, `cffi`, `rpds-py`) have prebuilt `win_arm64` wheels — see the [ARM64 install notes](#arm64-windows) below for the one extra `pip` flag you need.

#### Outlook "Programmatic Access" security prompts

When the MCP server first touches Outlook's COM API, Outlook may show a dialog: *"A program is trying to access email address information stored in Outlook"* (the Object Model Guard / Programmatic Access prompt). This is Outlook's protection against malicious automation.

You have three options:

1. **Click "Allow access for 10 minutes"** every time you start a session. Fine for casual use.
2. **Get the antivirus status to "Valid"** in *File > Options > Trust Center > Trust Center Settings > Programmatic Access*. When that line reads `Valid`, the "Never warn me about suspicious activity" radio becomes selectable and the prompts go away. On most personal machines with current Defender this works out of the box.
3. **Apply the registry policy** in [`docs/suppress-outlook-oom-prompts.reg`](docs/suppress-outlook-oom-prompts.reg). From an **elevated** PowerShell or Command Prompt (Win+X → *Terminal (Admin)*), run:

   ```powershell
   reg import "C:\path\to\outlook-desktop-mcp\docs\suppress-outlook-oom-prompts.reg"
   ```

   Then fully quit Outlook (check Task Manager for stray `OUTLOOK.EXE` processes) and reopen it. This writes `AdminSecurityMode=3` and approves all `PromptOOM*` categories under `HKLM\Software\Policies\Microsoft\Office\16.0\Outlook\Security`, which Outlook honors regardless of AV status. On Intune/MDM-managed corporate devices, `HKCU\Software\Policies\...\Outlook` is locked and the HKLM keys may be overwritten on next policy sync — if `reg import` fails or the prompts come back, ask IT to push the equivalent settings via Group Policy.

> **Windows 11 ARM64 note:** Defender does not register with Outlook's `IOfficeAntiVirus` interface on ARM64, so Trust Center shows *"Antivirus status: Invalid"* and the *"Never warn me about suspicious activity"* radio stays greyed out **even when Outlook is launched as Administrator**. Option 2 is unavailable on ARM64; the `reg import` from option 3 is the only durable suppression path.

### macOS

- **Microsoft Outlook for Mac** — version 16.x or later
- **Python 3.12+**
- **Outlook must be running** when the MCP server starts

#### Required macOS permissions

The first time a tool runs, macOS will show **two permission prompts** that you must approve:

1. **Privacy & Automation** — a system dialog asks: *"python3.12 wants to control Microsoft Outlook"*. Click **Allow** to let the server send AppleScript commands to Outlook.

2. **Accessibility** — to read your Exchange/M365 inbox, the server uses macOS UI scripting (System Events). This requires Accessibility access for `python3.12`:
   - Open **System Settings > Privacy & Security > Accessibility**
   - Find **python3.12** in the list (it appears after the first prompt)
   - Toggle it **on**

   Without Accessibility enabled, calendar, tasks, and local folder tools will work, but listing Exchange inbox messages will return empty results.

Both permissions are one-time setup — macOS remembers them for future sessions.

## Available Tools by Platform

### Email

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `send_email` | yes | yes | Send an email with To/CC/BCC, plain text or HTML body, and file attachments |
| `create_draft` | yes | yes | Compose an email and save it to Drafts without sending |
| `list_emails` | yes | yes | List recent emails from any folder, with optional unread filter |
| `read_email` | yes | yes | Read full email content by entry ID or subject search |
| `search_emails` | yes | yes | Full-text search across email subjects and bodies |
| `reply_email` | yes | yes | Reply or reply-all, preserving the conversation thread, with attachments |
| `forward_email` | yes | yes | Forward a message to new recipients with an optional comment |
| `mark_as_read` | yes | yes | Mark a specific email as read |
| `mark_as_unread` | yes | yes | Mark a specific email as unread |
| `move_email` | yes | yes | Move an email to Archive, Trash, or any folder |
| `snooze_email` | no | yes | Postpone an email: follow-up flag + Outlook reminder at a chosen time, optionally moving it to a snooze folder |
| `unsnooze_email` | no | yes | Clear the follow-up flag and reminder, optionally moving the email back to the inbox |
| `list_folders` | yes | yes | Browse the folder hierarchy with item counts (on macOS also each folder's id, path, depth and account) |

### Calendar

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_calendars` | no | yes | List every calendar across all accounts, with id, account and event count |
| `list_events` | yes | yes | List events within a date range, optionally scoped to one calendar |
| `get_event` | yes | yes | Read full event details by entry ID |
| `create_event` | yes | yes | Create a personal calendar appointment, optionally in a chosen calendar |
| `create_meeting` | yes | yes | Create a meeting and send invitations to attendees |
| `update_event` | yes | yes | Modify an existing event's subject, time, location, etc. |
| `delete_event` | yes | yes | Delete an appointment or cancel a meeting |
| `respond_to_meeting` | yes | yes | Accept, decline, or tentatively accept a meeting invite |
| `move_event` | no | yes | Move an event to another calendar — refuses any event that has attendees |
| `search_events` | yes | yes | Search calendar events by keyword, optionally scoped to one calendar |

> **macOS note — multiple calendars:** an Outlook profile usually holds many
> calendars spread over several accounts, and their names are **not unique**
> (two accounts can each have a "Calendar"). Call `list_calendars` and pass the
> numeric `calendar_id` to `list_events`, `search_events` and `create_event`;
> every event returned also reports the `calendar` it lives in. Omitting
> `calendar_id` spans every calendar of every account, which on a large profile
> can take tens of seconds.
>
> **macOS note — moving events:** Outlook for Mac cannot reassign an event's
> calendar in place (`set calendar of <event>` fails), so `move_event` copies
> the event and deletes the original. Because deleting a meeting cancels it for
> everyone invited and the copy then re-invites them, **`move_event` refuses any
> event with attendees** and there is no override flag. The copy preserves
> subject, times, location, body, organizer, categories and recurrence; the
> original is deleted only after the copy is found and verified, and the event's
> `entry_id` changes as a result.
>
> **macOS note — recurring events:** the date range is matched against each
> event's own start time, which for a recurring series is its *first*
> occurrence. An ongoing series that began before `start_date` is therefore not
> listed even though it still has occurrences in the range. Events that do
> appear are flagged with `is_recurring`.
>
> **macOS note:** `respond_to_meeting` acts on the meeting **invite message** in your mailbox, so its `entry_id` is that message's id (not a calendar event id). It also accepts `send_response` (default true) and an optional `comment` to the organizer.
>
> **macOS note:** `create_meeting` sends the invitations with Outlook's `send meeting` command. Pass `send_invites=false` to save the meeting with its attendees without notifying anyone, then review and send it from Outlook.

### Tasks

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_tasks` | yes | yes | List pending or completed tasks, sorted by due date |
| `get_task` | yes | yes | Read full task details including body and completion status |
| `create_task` | yes | yes | Create a new task with subject, due date, importance |
| `update_task` | yes | yes | Modify an existing task's subject, body, due date, importance, or completion |
| `search_tasks` | yes | yes | Search tasks by keyword in their title |
| `complete_task` | yes | yes | Mark a task as complete |
| `delete_task` | yes | yes | Remove a task |

### Attachments

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_attachments` | yes | yes | List all attachments on an email or calendar event |
| `save_attachment` | yes | yes | Download an attachment to a local directory |

### Categories

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_categories` | yes | yes | List all available color categories in Outlook |
| `set_category` | yes | yes | Set or clear categories on an email, event, or task |

> **macOS note:** `set_category` requires an `item_type` argument (`"email"`, `"task"`, or `"event"`). Unlike Windows EntryIDs, macOS numeric IDs are only unique *within* an item type, so the type is needed to target the correct item. On Windows the globally-unique EntryID makes `item_type` unnecessary.

### Rules, Out of Office (Windows only)

These tools rely on COM-specific APIs (the Rules object model and MAPI property accessors) that Outlook for Mac does not expose through AppleScript.

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_rules` | yes | — | List all mail rules with enabled/disabled status |
| `toggle_rule` | yes | — | Enable or disable a mail rule by name |
| `get_out_of_office` | yes | — | Check whether Out of Office auto-reply is on or off |

### Accounts (Windows only)

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_accounts` | yes | — | List the Outlook accounts (stores) in the profile |

> **Windows note:** most tools accept an optional `account` argument (the
> account's display name or a substring of it) to act on a store other than the
> primary one. Use `list_accounts` to see the names. On macOS, calendars are
> addressed with `calendar_id` from `list_calendars` instead.

**Total: 34 tools on each platform — 30 shared, 4 Windows-only, 4 macOS-only.**

### Results and errors (macOS)

Every macOS tool returns JSON. Successful calls carry a `status` (such as
`sent`, `created`, `moved`) and the item's `entry_id`; failures return
`{"error": "...", "code": "..."}`, where `code` is one of `invalid_argument`,
`not_found`, `permission_denied`, `timeout`, `outlook_not_running`,
`applescript_error`, `parse_error` or `unexpected`. Dates are returned as ISO
8601 regardless of the Mac's language. A `timeout` on a tool that changes
something means the change may still have gone through, so check before
retrying.

## Architecture Details

### Windows: COM Bridge (`com_bridge.py`)

All Outlook COM operations run on a dedicated thread using the Single-Threaded Apartment (STA) model, as required by COM. The async MCP event loop dispatches tool calls to this thread via a queue and awaits results, keeping COM threading rules respected and the MCP protocol non-blocking.

```
MCP tool call (async)
  → bridge.call(func, args)
    → queued to STA thread
      → func(outlook, namespace, args) executes on COM thread
    → result returned via threading.Event
  → JSON response back to MCP client
```

Each tool's inner function receives the live `Outlook.Application` and `MAPI.Namespace` COM objects and works directly with the Outlook Object Model — `GetItemFromID`, `CreateItem`, `Items.Restrict` with DASL filters, and so on.

### macOS: AppleScript Bridge (`applescript_bridge.py`)

Each tool call builds an AppleScript string and executes it as a subprocess via `osascript`. There is no persistent connection — every call is stateless.

```
MCP tool call (async)
  → build AppleScript string
    → asyncio.create_subprocess_exec("osascript", "-e", script)
    → parse stdout text into structured data
  → JSON response back to MCP client
```

Each tool constructs a single AppleScript that fetches all needed data in one `osascript` call (no per-message subprocess loops). Results come back as delimited text, which the server parses into the same JSON structure the Windows server produces.

**Key differences from Windows:**

- Entry IDs on macOS are **numeric** (e.g. `42`), not hex strings. They identify items within their folder context.
- Folder references use AppleScript's **locale-independent keywords** (`inbox`, `sent items`, `drafts`, `deleted items`) rather than localized folder names.
- Search uses AppleScript's `whose` clause (e.g. `messages whose subject contains "query"`) instead of DASL filters.
- User input is escaped for safe embedding in AppleScript strings to prevent script injection.

## Install from Source

### Windows (x64)

```bash
git clone https://github.com/Aanerud/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python -m venv .venv
.venv\Scripts\activate
pip install -e .
python .venv\Scripts\pywin32_postinstall.py -install
```

Register from source using the launcher script:

```bash
claude mcp add outlook-desktop -- powershell.exe -Command "& 'C:\path\to\outlook-desktop-mcp\outlook-desktop-mcp.cmd' mcp"
```

### Windows (ARM64)

Some transitive dependencies may resolve to versions without a `win_arm64` wheel, which then fail to build because they need a Rust toolchain plus OpenSSL. Force wheels-only resolution:

```powershell
git clone https://github.com/Aanerud/outlook-desktop-mcp.git
cd outlook-desktop-mcp
& "C:\Program Files\Python312-arm64\python.exe" -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install --only-binary=:all: pywin32 mcp
python -m pip install --no-deps -e .
python .venv\Scripts\pywin32_postinstall.py -install
```

The package depends on the base `mcp` package only. The `[cli]` extra of `mcp` (which pulls in `cryptography`) is needed just for the `mcp` developer tools (`mcp dev`, `mcp inspector`); install it with `pip install -e ".[cli]"` if you want them.

Register from source the same way as x64:

```bash
claude mcp add outlook-desktop -- powershell.exe -Command "& 'C:\path\to\outlook-desktop-mcp\outlook-desktop-mcp.cmd' mcp"
```

### macOS

```bash
git clone https://github.com/Aanerud/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
```

Register from source:

```bash
claude mcp add outlook-desktop -- /path/to/outlook-desktop-mcp/.venv/bin/python -m outlook_desktop_mcp
```

## Usage Examples

Once registered, just talk to Claude naturally:

- *"Show me my 10 most recent inbox emails"*
- *"Read the email from Taylor about MLADS"*
- *"Send an email to alice@example.com about the project update"*
- *"What's on my calendar this week?"*
- *"Create a meeting with bob@example.com tomorrow at 2pm for 30 minutes"*
- *"Save the attachment from that email to my Downloads folder"*
- *"Create a task to review the quarterly report, due Friday, high importance"*
- *"Mark that email as read and move it to archive"*

- *"What categories do I have? Set this email to 'Follow-up'"*

Windows-only examples:

- *"List my mail rules"*
- *"Am I set as Out of Office?"*

## Why Not Microsoft Graph?

| | Microsoft Graph | outlook-desktop-mcp |
|---|---|---|
| Entra app registration | Required | Not needed |
| Admin consent | Required for mail permissions | Not needed |
| OAuth token management | You handle refresh tokens | Not needed |
| Tenant configuration | Required | Not needed |
| Works offline / cached | No | Yes (reads from local cache) |
| Setup time | 30-60 minutes | 2 minutes |
| Auth requirement | **Your own OAuth flow** | **Outlook is open** |

## Project Structure

```
outlook-desktop-mcp/
  src/outlook_desktop_mcp/
    entrypoint.py            # Platform detection → routes to correct server
    server.py                # Windows MCP server (34 tools, COM automation)
    server_mac.py            # macOS MCP server (34 tools, AppleScript)
    com_bridge.py            # Async-to-COM threading bridge (Windows)
    applescript_bridge.py    # Async osascript execution (macOS)
    tools/
      _folder_constants.py   # Outlook enums and constants (Windows)
    utils/
      formatting.py          # Email/event/task data extraction (Windows)
      errors.py              # COM error formatting (Windows)
      com_dates.py           # Date parsing and task reminders (Windows)
      applescript_helpers.py # AppleScript escaping, date building (macOS)
  tests/
    test_*.py                # Unit tests (pytest) — no Outlook needed, run in CI
    integration/             # Scripts that drive a live Windows Outlook
      phase1_com_test.py     # Email COM validation
      phase3_mcp_test.py     # Email MCP test
      calendar_com_test.py   # Calendar COM validation
      calendar_mcp_test.py   # Calendar MCP test
      extras_com_test.py     # Tasks/attachments/categories/rules/OOF COM test
      extras_mcp_test.py     # Tasks/attachments/categories/rules/OOF MCP test
  outlook-desktop-mcp.cmd   # Windows launcher script
  pyproject.toml
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for the branching strategy and development setup.

## License

See [LICENSE](LICENSE) file.
