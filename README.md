# pyxll-claude

A [PyXLL](https://www.pyxll.com) extension that embeds a Claude Code terminal directly into Microsoft Excel as a dockable Custom Task Pane.

Chat with Claude in natural language, ask it to write Python functions, and register those functions with Excel immediately — no restarts or manual config edits required.

---

## How it works

1. Configure a workspace folder in `pyxll.cfg`.
2. Click **Claude** in the PyXLL ribbon tab to open the task pane.
3. Chat with Claude — ask it to build formulas, explain data, or write Python functions.
4. Claude writes functions into `pyxll_claude_functions.py` in your workspace.
5. Click **Load Functions** in the Functions tab to register them with Excel instantly.

The workspace is automatically bootstrapped with a `CLAUDE.md` and a `/fetch-pyxll-docs` skill so Claude has full PyXLL context before you type anything. Authentication and model selection are handled entirely by the `claude` CLI.

---

## Requirements

- Windows (Excel + PyXLL are Windows-only)
- Python 3.10+ (must match the Python that PyXLL uses)
- [PyXLL](https://www.pyxll.com) 5.1+ installed and configured in Excel
- [Claude Code CLI](https://claude.ai/code) (`claude`) installed and authenticated
- [uv](https://docs.astral.sh/uv/) (for development)

---

## Installation

```bash
uv pip install pyxll-claude
```

Or for development:

```bash
uv sync
uv pip install -e .
```

PyXLL discovers the extension automatically via setuptools entry points — no changes to `pyxll.cfg` are needed beyond the workspace setting below.

---

## Configuration

Add a `[CLAUDE]` section to your `pyxll.cfg`:

```ini
[CLAUDE]
workspace = C:\path\to\your\workspace
```

The workspace folder is created and bootstrapped automatically on first open. It will contain:

```
pyxll_claude_functions.py          ← your Excel functions (Claude edits this)
CLAUDE.md                          ← PyXLL context loaded automatically by Claude
.claude/
  settings.local.json              ← pre-approved permissions for pyxll.com
  skills/
    fetch-pyxll-docs/
      SKILL.md                     ← skill that fetches authoritative PyXLL docs
```

---

## Usage

### Writing functions

Ask Claude to write a function in plain English:

> *"Write an Excel function that returns the transpose of a range as a numpy array"*

Claude fetches the live PyXLL documentation, writes the function into `pyxll_claude_functions.py`, then you click **Load Functions** to register it.

Check the PyXLL log file if something doesn't appear.

---

## Project structure

```
pyxll-claude/
├── pyproject.toml
├── docs/
│   └── spec.md                     # Full project specification
└── src/
    └── pyxll_claude/
        ├── __init__.py             # pyxll_modules() and pyxll_ribbon() entry points
        ├── task_pane.py            # PySide6 CTP widget (Terminal + Functions tabs)
        ├── terminal.py             # pywinpty PTY wrapper and I/O threads
        ├── workspace.py            # First-run workspace initialisation
        ├── pyxll_claude_functions.py         # @xl_func ribbon callbacks and @xl_on_open handler
        └── resources/
            ├── ribbon.xml          # Excel ribbon XML
            └── terminal.html       # xterm.js page loaded by QWebEngineView
```

---

## Architecture notes

- **Terminal:** xterm.js inside a `QWebEngineView`, connected to a pywinpty ConPTY. `windowsMode: true` is set to prevent resize conflicts between xterm.js and ConPTY.
- **Function loading:** `pyxll_claude_functions.py` is imported/reloaded and `pyxll.rebind()` is called, which re-scans all modules for `@xl_func` callables and registers them with Excel immediately.
- **Workspace bootstrap:** `workspace.py` creates missing files on first open but never overwrites existing ones, so user edits are preserved across upgrades.
