# pyxll-claude — Project Specification

## 1. Overview

`pyxll-claude` is a PyXLL extension that embeds a Claude Code terminal directly
into Microsoft Excel as a dockable Custom Task Pane. The terminal runs the `claude`
CLI as a subprocess inside a user-configured **workspace folder**, which is
automatically bootstrapped with a `CLAUDE.md` and a `/fetch-pyxll-docs` skill so
Claude has full PyXLL context before the user types anything.

Primary workflow:

1. Configure a workspace folder in `pyxll.cfg` under `[CLAUDE] workspace`.
2. Open Excel and click the **Claude** button in the PyXLL ribbon tab.
3. Chat with Claude in natural language — ask it to build formulas, explain data,
   or write Python functions.
4. Claude writes or edits functions in `pyxll_claude_functions.py` in the workspace
   and calls `pyxll_reload` via MCP to register them with Excel immediately.
5. Functions are live in Excel with no manual steps required.

---

## 2. Goals

- Provide an AI coding assistant that understands the Excel/PyXLL context.
- Allow non-developers to create custom Excel functions through conversation.
- Keep setup minimal: install the package, add one `pyxll.cfg` entry, done.
- Leverage the `claude` CLI rather than the raw Anthropic API so authentication,
  model selection, and tool use are all handled by Claude Code itself.

---

## 3. Feature Phases

### Phase 1 — Claude Terminal (MVP) ✓

**Scope:**
- A Custom Task Pane containing an embedded xterm.js terminal.
- The user configures a workspace directory in `pyxll.cfg` under `[CLAUDE] workspace`.
- If the setting is missing or the folder does not exist, the terminal displays a
  clear error message with setup instructions instead of starting claude.
- On open, a `claude` subprocess is started **in the workspace folder** so Claude
  Code picks up the workspace's `CLAUDE.md` and `.claude/skills/fetch-pyxll-docs/`
  automatically, giving Claude PyXLL context before the user types anything.
- The terminal relays input/output between the user and the `claude` process with
  full ANSI colour support via xterm.js.
- Authentication and model selection are delegated entirely to the `claude` CLI.

**Implementation notes:**
- Terminal is rendered by **xterm.js** inside a `QWebEngineView`. `windowsMode: true`
  is required to prevent xterm.js and ConPTY from both reflowing lines on resize.
- PTY is provided by **pywinpty** (`winpty.PTY`, ConPTY backend). A dedicated
  background read thread drains PTY output and emits Qt signals; a write queue
  serialises stdin writes from the main thread.
- `claude` is resolved via `shutil.which()`; `.cmd`/`.bat` entries are routed
  through `cmd.exe /d /c`.
- Resize: a 150 ms debounced `ResizeObserver` fires `fitAddon.fit()` and
  `bridge.resizeTerminal()` in the same callback so xterm.js and the PTY always
  change size together.

**Out of scope in Phase 1:** function registration, Excel workbook awareness.

---

### Phase 2 — Workspace & Function Loading ✓

**Scope:**
- The workspace folder is configured in `pyxll.cfg`:
  ```ini
  [CLAUDE]
  workspace = C:\path\to\your\workspace
  ```
- On first open, the workspace is automatically bootstrapped with:
  - `CLAUDE.md` — workspace-specific PyXLL context for Claude.
  - `.claude/skills/fetch-pyxll-docs/SKILL.md` — the `/fetch-pyxll-docs` skill.
  - `.claude/skills/fetch-pyxll-docs/scripts/search-pyxll-docs.sh` — keyword search script.
  - `.claude/settings.local.json` — pre-approved permissions.
  - `pyxll_claude_functions.py` — an empty module stub ready for `@xl_func` functions.
- The CTP gains a second tab, **Functions**, alongside the existing **Terminal** tab.
- The Functions tab shows a read-only viewer for `pyxll_claude_functions.py` that
  auto-refreshes via `QFileSystemWatcher` whenever Claude edits the file.
- A **Load Functions** button imports or reloads `pyxll_claude_functions` from the
  workspace and calls `pyxll.rebind()` to register any `@xl_func`-decorated functions
  with Excel immediately. Success and error states are shown inline below the button.

**Workflow:**
```
User asks Claude to write a function
  → Claude edits pyxll_claude_functions.py in the workspace
    → Functions tab auto-refreshes (QFileSystemWatcher)
      → User clicks Load Functions (manual fallback)
        → importlib.reload("pyxll_claude_functions")
          → pyxll.rebind()
            → Function appears in Excel
```

**Out of scope in Phase 2:** function listing/management UI, delete/edit actions.

---

### Phase 3 — Autonomous Reload & Excel Control via MCP ✓

**Scope:**
- A minimal MCP (Model Context Protocol) server over HTTP/SSE runs inside Excel's
  Python process, bound to `127.0.0.1` (localhost only) on the first free port
  found in a configurable range (default `54717–54816`).
- The server and the claude subprocess are started on a background thread to avoid
  blocking Excel's main thread with socket operations.
- The actual port is written to `.mcp.json` in the workspace root before claude
  starts. Claude Code reads this file automatically at startup.
- `settings.local.json` is extended with pre-approved MCP tool permissions
  (`mcp__pyxll__*`) and `enabledMcpjsonServers: ["pyxll"]` so the server is
  trusted without any user prompt.
- The workspace `CLAUDE.md` is updated: claude is instructed to call `pyxll_reload`
  after every edit and loop until it returns `OK`, eliminating the need for the
  user to click **Load Functions**.
- The **Load Functions** button is **kept** as a manual fallback for edits made
  outside the claude session.

**Workflow:**
```
Claude edits pyxll_claude_functions.py
  → calls pyxll_reload MCP tool
    → schedule_call → Excel main thread: importlib.reload + pyxll.rebind
      → returns "OK" or Python traceback
        → if error: Claude fixes the file and calls pyxll_reload again
          → repeats until OK → function immediately available in Excel
```

**MCP tools:**

| Tool | Inputs | Description |
|------|--------|-------------|
| `pyxll_reload` | — | Reloads `pyxll_claude_functions.py` and calls `pyxll.rebind()`. Returns `"OK"` or a Python traceback. |
| `pyxll_get_log` | `lines?` (int, default 50) | Returns the last N lines from the PyXLL log file. Log path is read from `pyxll.get_config()` (`[LOG] path` + `[LOG] file`) at call time. |
| `pyxll_get_selection` | — | Returns the current selection address and sheet name. |
| `pyxll_read_range` | `ref` (str), `sheet?` (str) | Reads cell values from an Excel range via the COM API. Returns a 2-D list of rows. |
| `pyxll_write_range` | `ref` (str), `values` (2-D list), `sheet?` (str) | Writes values or formulas to an Excel range via the COM API. Uses `.Formula` for cells starting with `=`, `.Value2` otherwise. COM parameterised properties use their `Get<Name>` Python equivalents (e.g. `GetResize`, `GetOffset`). |

**Architecture:**
- **Transport:** HTTP/SSE — server runs in a background daemon thread inside the
  existing Excel Python process; no separate subprocess; bound to `127.0.0.1` only.
- **Dependencies:** stdlib only (`http.server.ThreadingHTTPServer`, `threading`,
  `queue`). No new packages.
- **Threading:** `pyxll_reload`, `pyxll_read_range`, `pyxll_write_range`, and
  `pyxll_get_selection` must run on Excel's main thread. The HTTP thread uses
  `pyxll.schedule_call()` to dispatch the work and `threading.Event` (30 s timeout)
  to wait for the result. `pyxll_get_log` is pure file I/O on the HTTP thread.

**Port selection and communication:**
```ini
[CLAUDE]
workspace      = C:\path\to\workspace
mcp_port_range = 54717-54816   ; optional, this is the default
```
Ports are tried in order via `socket.bind(('127.0.0.1', port))`; the first that
succeeds is used. The chosen port is written to `.mcp.json` in the workspace root:
```json
{
  "mcpServers": {
    "pyxll": { "type": "sse", "url": "http://localhost:54717/sse" }
  }
}
```
This file is overwritten each session. `settings.local.json` opts in via
`"enabledMcpjsonServers": ["pyxll"]` so no user confirmation prompt appears.

If all ports are exhausted, a warning is logged and claude starts without MCP;
the Load Functions button remains available.

**New files:**
- `src/pyxll_claude/mcp_server.py` — `PyXLLMCPServer` class and `find_free_port()`

**Modified files:**
- `src/pyxll_claude/task_pane.py` — background thread for MCP start, port scan,
  `.mcp.json` write, server lifecycle
- `src/pyxll_claude/workspace.py` — `_CLAUDE_MD` and `_SETTINGS_LOCAL_JSON`
  templates updated; MCP permissions and `enabledMcpjsonServers` added

**Out of scope in Phase 3:** real-time error streaming to the terminal;
function listing/management UI (Phase 4).

---

### Phase 4 — Function Management

**Scope:**
- A panel within the Functions tab listing all currently loaded Claude functions.
- Each entry shows: name, signature, description, source code.
- Actions: **Edit** (opens the relevant section in the terminal for Claude to revise),
  **Delete** (removes from module and calls `pyxll.rebind()`).
- Functions in `pyxll_claude_functions.py` automatically survive a PyXLL reload
  because the file is re-imported on startup.

---

## 4. Architecture

### 4.1 Package Layout

```
src/pyxll_claude/
├── __init__.py      # Entry points: pyxll_modules(), pyxll_ribbon()
├── task_pane.py     # CTP widget, TerminalBridge, show_claude_pane()
├── terminal.py      # ClaudeProcess — pywinpty PTY wrapper + I/O threads
├── workspace.py     # First-run workspace initialisation
├── mcp_server.py    # PyXLLMCPServer — HTTP/SSE MCP server + tools
├── xl_functions.py  # PyXLL ribbon callbacks and @xl_on_open handler
└── resources/
    ├── __init__.py
    ├── ribbon.xml
    ├── claude.png      # Ribbon button icon
    └── terminal.html   # xterm.js page loaded by QWebEngineView
```

### 4.2 Entry Points

| Entry point | Function | Purpose |
|-------------|----------|---------|
| `pyxll.modules` | `pyxll_claude:pyxll_modules` | Returns `["pyxll_claude.xl_functions"]` |
| `pyxll.ribbon`  | `pyxll_claude:pyxll_ribbon`  | Returns ribbon XML for the Claude button |

### 4.3 Custom Task Pane

Uses **PySide6** (Qt6) via PyXLL's `create_ctp()`. The CTP is a `QTabWidget`
with two tabs:

| Tab | Contents |
|-----|----------|
| Terminal  | `QWebEngineView` hosting the xterm.js terminal |
| Functions | Read-only `QTextEdit` viewer for `pyxll_claude_functions.py` + **Load Functions** button |

A new `ClaudeTerminalWidget` is created on every `show_claude_pane()` call; the
previous widget is destroyed by PyXLL when the pane is closed.

### 4.4 Claude CLI Terminal

```
pywinpty PTY (ConPTY)  ←→  read thread → data_received signal → TerminalBridge
                             write queue ← TerminalBridge ← QWebChannel ← xterm.js
```

**Workspace selection:** `[CLAUDE] workspace` from `pyxll.cfg`, validated and
bootstrapped before the PTY is started. Claude Code auto-loads `CLAUDE.md` and
`.claude/skills/` from the workspace cwd.

**xterm.js / ConPTY compatibility:** `windowsMode: true` disables xterm.js's
internal line-reflow so it does not conflict with ConPTY's own reflow on resize.

**Resize:** a single 150 ms `ResizeObserver` callback calls `fitAddon.fit()` then
`bridge.resizeTerminal(cols, rows)` so xterm.js and the PTY dimensions always
change in the same tick.

### 4.5 Workspace Initialisation

`workspace.ensure_workspace_initialized(path)` is called once each session before
starting the PTY. It creates missing files but never overwrites existing ones:

| File | Purpose |
|------|---------|
| `CLAUDE.md` | Instructs Claude to use `pyxll_reload` after edits and reach for MCP tools proactively |
| `.claude/skills/fetch-pyxll-docs/SKILL.md` | Skill: fetches PyXLL docs index, individual pages, or searches via script |
| `.claude/skills/fetch-pyxll-docs/scripts/search-pyxll-docs.sh` | Caches and keyword-searches `llms-full.txt` locally |
| `.claude/settings.local.json` | Pre-approved permissions (bash, curl, file I/O, MCP tools) + `enabledMcpjsonServers` |
| `pyxll_claude_functions.py` | Stub module with `from pyxll import xl_func` |

### 4.6 Function Loading Flow

**Primary path (Phase 3 — via MCP):**
```
Claude edits pyxll_claude_functions.py
  → calls pyxll_reload MCP tool
    → HTTP thread: pyxll.schedule_call(_do_reload)
      → Excel main thread: importlib.reload("pyxll_claude_functions")
        → pyxll.rebind()
          → functions registered in Excel
            → "OK" returned to Claude
```

**Fallback path — Load Functions button:**
```
User clicks Load Functions
  → workspace added to sys.path (once)
    → importlib.reload / import_module("pyxll_claude_functions")
      → pyxll.rebind()
        → functions registered in Excel
```

`pyxll.rebind()` re-scans all modules known to PyXLL, so any `@xl_func`-decorated
callable in `pyxll_claude_functions` is picked up without restarting Excel.

---

## 5. Technical Decisions

| Decision | Choice | Rationale |
|----------|--------|-----------|
| UI framework | PySide6 | Active Qt6 bindings; first-class PyXLL CTP support |
| Claude backend | `claude` CLI via ConPTY | Delegates auth, model selection, and tool use to Claude Code |
| PTY | pywinpty `PTY` class (ConPTY) | Native Windows ConPTY; better than pipe-based `QProcess` for interactive TUIs |
| Terminal rendering | xterm.js in `QWebEngineView` | Full ANSI/VT emulation; `windowsMode: true` resolves ConPTY resize conflicts |
| Workspace folder | User-configured in `pyxll.cfg [CLAUDE]` | Separates user code from the extension; survives package upgrades |
| First-run bootstrap | `workspace.py` creates files if absent | Zero manual setup beyond the single `pyxll.cfg` entry |
| Function registration | MCP `pyxll_reload` tool (Phase 3) / Load Functions button (fallback) | Claude iterates autonomously; button retained for manual edits |
| Streaming | Background read thread + `data_received` signal | Non-blocking; Qt signal delivery marshals bytes to main thread safely |
| Dependency management | `uv` | Fast, reproducible |

---

## 6. Non-Goals

- Direct Anthropic API or SDK usage — the `claude` CLI handles this.
- API key management within this package.
- Supporting Excel versions older than 2016 (no Custom Task Pane API).
- Multi-user or cloud sync of registered functions.
- A standalone (non-Excel) mode.
