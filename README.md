# pyxll-claude

A [PyXLL](https://www.pyxll.com) extension that embeds a Claude Code terminal directly into Microsoft Excel as a dockable Custom Task Pane.

Chat with Claude in natural language, ask it to write Python functions, and have those functions registered with Excel automatically — no restarts or manual config edits required.

---

## How it works

1. Click **Claude** in the PyXLL ribbon tab to open the task pane.
2. Chat with Claude — ask it to build formulas, explain data, or write Python functions.
3. Claude writes functions into `pyxll_claude_functions.py` in your workspace.
4. Claude calls the built-in `pyxll_reload_module` MCP tool to register the new functions with Excel immediately.

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

PyXLL discovers the extension automatically via setuptools entry points — no changes to `pyxll.cfg` are required.

---

## Configuration

The workspace defaults to a `pyxll_claude_workspace` folder next to the PyXLL add-in. To use a different location, add a `[CLAUDE]` section to your `pyxll.cfg`:

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

Claude fetches the live PyXLL documentation, writes the function into `pyxll_claude_functions.py`, then calls `pyxll_reload_module` via MCP to register it with Excel — no button press needed.

You can also view and edit `pyxll_claude_functions.py` directly in the **Functions** tab, which hosts a Monaco code editor. Saving the file there (Ctrl+S) triggers the same reload automatically.

Check `%APPDATA%\PyXLL\pyxll.log` if a function doesn't appear.

---

## Project structure

```
pyxll-claude/
├── pyproject.toml
├── docs/
│   └── spec.md                        # Full project specification
└── src/
    └── pyxll_claude/
        ├── __init__.py                # pyxll_modules() and pyxll_ribbon() entry points
        ├── task_pane.py               # PySide6 CTP widget (Terminal + Functions tabs)
        ├── mcp_server.py              # Local MCP server (pyxll_reload_module tool)
        ├── workspace.py               # First-run workspace initialisation
        ├── xl_functions.py            # @xl_on_open handler and ribbon callbacks
        ├── widgets/
        │   ├── terminal.py            # Claude CLI subprocess + xterm.js terminal
        │   └── editor.py              # Monaco editor for pyxll_claude_functions.py
        ├── webview/
        │   ├── client.py              # Parent-process WebViewClient (child-process embedding)
        │   └── child_process.py       # Child process hosting QWebEngineView
        └── resources/
            ├── ribbon.xml             # Excel ribbon XML
            ├── terminal.html          # xterm.js page
            └── editor.html            # Monaco editor page
```

---

## Architecture notes

- **Terminal:** xterm.js served from a child process `QWebEngineView`, connected to the `claude` CLI via a pywinpty ConPTY. The child process is embedded into the parent window via Win32 `SetParent`. `windowsMode: true` prevents resize conflicts between xterm.js and ConPTY.
- **Code editor:** Monaco editor also runs in a child-process `QWebEngineView`. It watches `pyxll_claude_functions.py` for external changes and reloads automatically.
- **MCP server:** A lightweight HTTP/SSE server runs on localhost and exposes tools that give Claude live access to Excel and the PyXLL environment. When Claude saves a function it calls `pyxll_reload_module` to reload the module and call `pyxll.rebind()`, registering new `@xl_func` callables with Excel on the spot. The full tool set is:

  | Tool | Purpose |
  |---|---|
  | `pyxll_reload_module` | Reload a workspace module and rebind `@xl_func` functions |
  | `pyxll_get_log` | Read the tail of the PyXLL log file |
  | `pyxll_get_config` | Read the parsed `pyxll.cfg` |
  | `pyxll_get_config_path` | Locate the `pyxll.cfg` file |
  | `pyxll_get_version` | PyXLL version string |
  | `pyxll_get_python_version` | Python version string |
  | `pyxll_get_selection` | Current Excel selection (sheet + address) |
  | `pyxll_read_range` | Read cell values from a worksheet range |
  | `pyxll_read_formulas` | Read cell formulas from a worksheet range |
  | `pyxll_write_range` | Write values or formulas to a worksheet range |
- **Workspace bootstrap:** `workspace.py` creates missing files on first open but never overwrites existing ones, so user edits are preserved across upgrades.
