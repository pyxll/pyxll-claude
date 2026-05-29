# pyxll-claude

A PyXLL extension that embeds a Claude Code terminal into Microsoft Excel as a
dockable Custom Task Pane. The terminal runs the `claude` CLI as a subprocess,
started from this project folder so it automatically picks up `CLAUDE.md` and
the `/fetch-pyxll-docs` skill — giving Claude full PyXLL context before the user
types anything. Users can chat with Claude, request Python functions, and have
those functions registered directly with Excel — no manual config edits.

---

## Project Setup

### Prerequisites

- Python 3.10+ (must match the Python that PyXLL uses)
- [uv](https://docs.astral.sh/uv/) installed
- PyXLL installed and configured in Excel (see https://www.pyxll.com/docs)

### Install for development

```bash
uv sync
uv pip install -e .
```

After the editable install, PyXLL discovers this package automatically via
setuptools entry points — no changes to `pyxll.cfg` are needed.

### Reload in Excel

After code changes, press **Reload PyXLL** in the PyXLL ribbon tab. This
re-imports all modules without restarting Excel.

---

## How PyXLL Loads This Extension

Entry points are declared in `pyproject.toml`:

```toml
[project.entry-points."pyxll"]
modules = "pyxll_claude:pyxll_modules"
ribbon  = "pyxll_claude:pyxll_ribbon"
```

- **`modules`** — `pyxll_modules()` in `src/pyxll_claude/__init__.py` returns the
  list of module names PyXLL should import and scan for decorators.
- **`ribbon`** — `pyxll_ribbon()` returns the XML loaded from
  `src/pyxll_claude/resources/ribbon.xml`, adding the Claude button to the ribbon.

No edits to `pyxll.cfg` are required.

---

## Claude CLI Subprocess

The task pane works by launching `claude` (the Claude Code CLI) as a child process
via `QProcess`, with the working directory set to this project's root folder.

```python
process = QProcess()
process.setWorkingDirectory(str(project_root))  # folder containing CLAUDE.md
process.start("claude", [])
```

Because the cwd contains `CLAUDE.md` and `.claude/skills/fetch-pyxll-docs/SKILL.md`,
Claude Code loads them automatically at startup — giving Claude full PyXLL context
before the user types anything. Authentication and model selection are handled
entirely by the `claude` CLI; this package does not manage API keys.

---

## PyXLL Documentation

Always fetch the authoritative docs before writing PyXLL-specific code:

```
/fetch-pyxll-docs
```

Or use the index directly: `https://www.pyxll.com/llms.txt`

---

## Key PyXLL Patterns

### Expose a function to Excel

```python
from pyxll import xl_func

@xl_func
def my_function(x: float, y: float) -> float:
    """Docstring appears as help text in Excel."""
    return x + y
```

### Add an Add-ins menu item

```python
from pyxll import xl_menu

@xl_menu("My Action", menu="My Extension")
def my_action():
    ...
```

### Run code when the add-in loads

```python
from pyxll import xl_on_open

@xl_on_open
def on_open(import_info):
    # import_info: list of (module_name, module, exc_info)
    ...
```

### Create a Custom Task Pane

```python
from pyxll import create_ctp, CTPDockPositionRight
from PySide6.QtWidgets import QWidget, QApplication

def show_my_pane():
    app = QApplication.instance() or QApplication([])
    widget = QWidget()
    create_ctp(widget, width=400, position=CTPDockPositionRight)
```

### Register functions dynamically at runtime

```python
from pyxll import rebind

# After exec()-ing new @xl_func decorated code into a module namespace:
rebind()
```

---

## Project Structure

```
pyxll-claude/
├── pyproject.toml                  # Package metadata, deps, entry points
├── CLAUDE.md                       # This file
├── .gitignore
├── docs/
│   └── spec.md                     # Full project specification
├── src/
│   └── pyxll_claude/
│       ├── __init__.py             # pyxll_modules() and pyxll_ribbon()
│       ├── task_pane.py            # PySide6 CTP widget
│       ├── terminal.py             # QProcess wrapper for the claude subprocess
│       ├── xl_functions.py         # @xl_func, @xl_menu, ribbon callbacks
│       └── resources/
│           ├── __init__.py
│           └── ribbon.xml          # Excel ribbon XML
└── .claude/
    └── skills/
        └── fetch-pyxll-docs/
                └── SKILL.md        # /fetch-pyxll-docs skill
```

---

## Testing

1. `uv pip install -e .` — install the editable package.
2. Open Excel (PyXLL must be installed in the same Python environment).
3. Confirm the **Claude AI** group appears in the PyXLL ribbon tab.
4. Click **Claude** to open the task pane.
5. After code changes, press **Reload PyXLL** in the ribbon.

Check `%APPDATA%\PyXLL\pyxll.log` for errors if something doesn't appear.
