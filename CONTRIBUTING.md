# Contributing to outlook-desktop-mcp

## Branching Strategy

```
feature/your-change → PR → preview → PR → main → tag vX.Y.Z → publish to PyPI
```

- **`main`** — stable. Never push directly. CI runs the unit tests on every push and PR.
- **`preview`** — integration testing branch. All PRs land here first.
- **Feature branches** — your working branches, created from `preview`.

## How to Contribute

1. **Fork** the repo to your GitHub account
2. **Clone** your fork locally
3. **Create a branch** from `preview`:
   ```bash
   git checkout preview
   git pull origin preview
   git checkout -b feature/my-change
   ```
4. **Make your changes**, commit, push to your fork
5. **Open a PR** into `preview` (not `main`)
6. Once reviewed and merged to `preview`, it will be tested there
7. Periodically, `preview` is merged into `main`. A release is cut by bumping `version` in
   `pyproject.toml` and pushing a matching tag (`git tag v0.5.0 && git push origin v0.5.0`);
   the publish workflow checks the tag against the version, runs the tests and uploads to PyPI.

## Development Setup

### Windows

Requires Outlook Desktop (Classic) running.

```bash
git clone https://github.com/YOUR-USERNAME/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python -m venv .venv
.venv\Scripts\activate
pip install -e ".[dev]"
python .venv\Scripts\pywin32_postinstall.py -install
```

### macOS

Requires Microsoft Outlook for Mac running. The first run asks for permission
to control Outlook (System Settings → Privacy & Security → Automation).

```bash
git clone https://github.com/YOUR-USERNAME/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python3 -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"
```

## Testing

The unit tests need no Outlook and run on any OS — this is what CI runs:

```bash
pytest
```

They replace the AppleScript bridge with a fake (see `tests/conftest.py`), so
they check the scripts each tool builds and how it parses Outlook's replies.
Tests for Windows helpers import only modules that do not need pywin32.

With Outlook open on Windows, the live scripts in `tests/integration/` exercise
the real COM and MCP layers:

```bash
.venv\Scripts\python tests\integration\phase1_com_test.py   # COM validation
.venv\Scripts\python tests\integration\phase3_mcp_test.py   # MCP protocol
```

On macOS, check a new AppleScript against the real dictionary without running
it by compiling it: `osacompile -o /tmp/x.scpt -e '<script>'` fails on unknown
terms.

## Adding New Tools

1. Implement the tool in `server.py` (Windows) and/or `server_mac.py` (macOS).
   - Windows: define the COM function (receives `outlook, namespace` as first
     args) and an `@mcp.tool()` handler that calls `bridge.call(...)`.
   - macOS: stack `@mcp.tool()` over `@_tool("<doing something>")` (add
     `mutates=True` if it changes anything) and build the AppleScript in the
     handler. Pass every id through `_item_id` / `_calendar_id`, every date
     through `_parse_iso`, and every string through `escape` — nothing the
     caller sends may reach the script unvalidated. Return JSON.
2. Write a detailed docstring — this is what LLMs see during tool discovery.
3. Add unit tests in `tests/test_*.py`.
4. Add the tool to the tables in `README.md` and update the tool counts.
5. Update the server's `instructions` string if it adds a capability category.
