# pyxll-claude

A [PyXLL](https://www.pyxll.com) extension that embeds a Claude Code terminal directly into Microsoft Excel as a dockable Custom Task Pane.

Chat with Claude in natural language to write Python functions, read and write cell data, format spreadsheets, manage sheets, and inspect the live state of your workbook — all without leaving Excel.

---

## How it works

1. Click **Claude** in the PyXLL ribbon tab to open the task pane.
2. Chat with Claude — ask it to build formulas, analyse data, format a table, or write Python functions.
3. Claude uses a set of built-in MCP tools to interact with Excel directly: reading ranges, writing values, formatting cells, managing sheets, and more.
4. When Claude writes new Python functions into `pyxll_claude_functions.py` it calls `pyxll_reload_module` to register them with Excel immediately — no restarts or button presses needed.

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
  settings.local.json              ← pre-approved permissions for safe tools
  skills/
    fetch-pyxll-docs/
      SKILL.md                     ← skill that fetches authoritative PyXLL docs
```

---

## Usage

### Writing Excel functions

Ask Claude to write a function in plain English:

> *"Write an Excel function that returns the transpose of a range as a numpy array"*

Claude fetches the live PyXLL documentation, writes the function into `pyxll_claude_functions.py`, then calls `pyxll_reload_module` to register it with Excel — no button press needed.

You can also view and edit `pyxll_claude_functions.py` directly in the **Functions** tab, which hosts a Monaco code editor. Saving the file there (Ctrl+S) triggers the same reload automatically.

### Working with spreadsheet data

Claude can read and write your spreadsheet directly:

> *"Read the sales table on Sheet2 and add a total row at the bottom"*
>
> *"Format the header row in A1:F1 as bold white text on a dark blue background"*
>
> *"Insert two rows above row 5 and add column headers"*
>
> *"Save the workbook"*

### Inspecting the workbook

> *"What sheets does this workbook have?"*
>
> *"What's in the current selection?"*
>
> *"Take a screenshot so I can see the current state"*
>
> *"Show me the named ranges defined in this workbook"*

Check `%APPDATA%\PyXLL\pyxll.log` if something doesn't appear as expected.

---

## MCP tools

A local MCP server runs inside Excel's Python process and exposes the following tools to Claude. Read-only tools are pre-approved and run without prompting; write tools ask for confirmation.

### PyXLL environment

| Tool | Purpose |
|---|---|
| `pyxll_reload_module` | Reload a workspace module and rebind `@xl_func` functions |
| `pyxll_get_log` | Read the tail of the PyXLL log file |
| `pyxll_get_config` | Read the parsed `pyxll.cfg` |
| `pyxll_get_config_path` | Locate the `pyxll.cfg` file |
| `pyxll_get_version` | PyXLL version string |
| `pyxll_get_python_version` | Python version string |

### Reading the workbook

| Tool | Purpose |
|---|---|
| `pyxll_get_sheets` | List all worksheet names in the active workbook |
| `pyxll_get_used_range` | Bounds and dimensions of the populated area of a sheet |
| `pyxll_get_selection` | Current Excel selection (sheet + address) |
| `pyxll_get_cell_info` | Value, formula, number format and full styling for one cell |
| `pyxll_get_named_ranges` | All defined names in the active workbook |
| `pyxll_read_range` | Read cell values from a worksheet range |
| `pyxll_read_formulas` | Read cell formulas from a worksheet range |
| `pyxll_screenshot` | Capture the Excel window as a PNG image |

### Writing and editing

| Tool | Purpose |
|---|---|
| `pyxll_write_range` | Write values or formulas to a worksheet range |
| `pyxll_clear_range` | Clear contents, formats, or both from a range |
| `pyxll_format_cells` | Apply font, fill, alignment, borders, and number format |
| `pyxll_auto_fit_column` | Auto-fit column width to content |
| `pyxll_set_column_width` | Set an explicit column width |
| `pyxll_set_row_height` | Set an explicit row height |

### Structure

| Tool | Purpose |
|---|---|
| `pyxll_insert_rows` | Insert blank rows, shifting existing rows down |
| `pyxll_insert_columns` | Insert blank columns, shifting existing columns right |
| `pyxll_merge_cells` | Merge a range into a single cell |
| `pyxll_unmerge_cells` | Unmerge previously merged cells |
| `pyxll_add_sheet` | Add a new worksheet to the workbook |
| `pyxll_name_range` | Create or update a workbook-level named range |
| `pyxll_save_workbook` | Save the active workbook |
| `pyxll_calculate` | Force recalculation at workbook, sheet, or range scope |

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
        ├── mcp_server/                # Local MCP server
        │   ├── server.py              # HTTP server and JSON-RPC dispatch
        │   └── tools/                 # One module per MCP tool
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
- **MCP server:** A lightweight Streamable HTTP server (MCP protocol 2025-03-26) runs on localhost in a daemon thread inside Excel's Python process. Each tool runs its COM operations on Excel's main thread via `pyxll.schedule_call`. The server survives PyXLL reloads by persisting as a singleton in `sys.modules`.
- **Workspace bootstrap:** `workspace.py` creates missing files on first open but never overwrites existing ones, so user edits are preserved across upgrades.
