# Contributing to outlook-desktop-mcp

## Branching Strategy

```
feature/your-change → PR → main → auto-publish to PyPI
```

- **`main`** — stable. Every push to `main` auto-publishes to PyPI, gated by the
  CI test job (lint + unit tests must pass before the publish step runs).
  Never push directly — always go through a PR.
- **Feature branches** — your working branches, created from `main`.

## How to Contribute

1. **Fork** the repo to your GitHub account
2. **Clone** your fork locally
3. **Create a branch** from `main`:
   ```bash
   git checkout main
   git pull origin main
   git checkout -b feature/my-change
   ```
4. **Make your changes**, commit, push to your fork
5. **Open a PR** into `main`. CI runs `ruff check` and the unit tests on every PR
6. Once reviewed and merged, the publish workflow releases to PyPI

## Development Setup

### Windows (requires Outlook Desktop Classic running)

```bash
git clone https://github.com/YOUR-USERNAME/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python -m venv .venv
.venv\Scripts\activate
pip install pywin32 "mcp[cli]" -e .[dev]
python .venv\Scripts\pywin32_postinstall.py -install
```

### macOS (requires Outlook for Mac running for the integration tests)

```bash
git clone https://github.com/YOUR-USERNAME/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python3 -m venv .venv
source .venv/bin/activate
pip install "mcp[cli]" -e .[dev]
```

## Testing

**Unit tests (any OS, no Outlook needed)** — these run in CI:

```bash
python tests/unit_mac_test.py
ruff check src tests
```

**Integration tests** require a live Outlook. On Windows, with Outlook Desktop
(Classic) open:

```bash
# COM validation (no MCP layer)
outlook-desktop-mcp.cmd test

# MCP protocol test
.venv\Scripts\python tests\phase3_mcp_test.py
```

## Adding New Tools

1. Define the COM function that does the work (receives `outlook, namespace` as first args)
2. Add an `@mcp.tool()` async handler in `server.py` that calls `bridge.call(your_function, ...)`
3. For macOS parity, add the matching handler in `server_mac.py`. Security rules
   for AppleScript construction: pass every user-supplied string through
   `escape(...)`, validate `entry_id` with `_validated_id(...)`, and clamp count
   arguments with `_clamp_count(...)`
4. Write a detailed docstring — this is what LLMs see during tool discovery
5. Add test coverage (mocked unit tests in `tests/unit_mac_test.py` where possible)
6. Update the `instructions` string if adding a new capability category
