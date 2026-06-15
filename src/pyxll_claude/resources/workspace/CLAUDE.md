# PyXLL Excel Functions Workspace

## Fetch PyXLL docs before writing or diagnosing any PyXLL code

Before writing, editing, reviewing, or diagnosing any issue with PyXLL code
in this workspace, fetch the authoritative PyXLL documentation:

```
/fetch-pyxll-docs
```

Your training-data knowledge of PyXLL is incomplete and may be wrong.
The live documentation is the only authoritative source for decorator
signatures, type annotations, and all other PyXLL-specific behaviour.
In particular, before suggesting a workaround for unexpected PyXLL behaviour,
check the docs first — PyXLL often has a built-in decorator parameter or
config setting that solves the problem directly.

Use `/fetch-pyxll-docs` for all PyXLL documentation lookups — including release
notes, changelog, and version history. Do not guess PyXLL documentation URLs directly.

## PyXLL MCP Tools

Use these tools proactively — do not ask the user to perform these actions manually:

- `pyxll_reload_module(module?)` — reload a single Python module (not the full add-in)
  and rebind @xl_func functions. Defaults to `pyxll_claude_functions` when omitted.
  Call this after every edit to `pyxll_claude_functions.py`. **Use this for all
  normal function edits** — it is safe and does not close the terminal.
- `pyxll_is_terminal_open()` — returns `True` if the Claude terminal task pane is
  currently open in Excel, `False` otherwise. Call this before `pyxll_reload_addin`
  to decide whether to warn the user.
- `pyxll_reload_addin()` — perform a full PyXLL add-in reload (equivalent to clicking
  "Reload PyXLL" in the ribbon). Closes ALL custom task panes including the Claude
  terminal. The MCP server itself survives the reload. Before calling: check
  `pyxll_is_terminal_open` — only warn the user that the terminal will close if it
  is actually open. Only use this when a full reload is genuinely necessary (e.g.
  after editing `pyxll.cfg` or adding a new module to the modules list).
- `pyxll_get_log` — whenever the user mentions the PyXLL log, asks about errors or
  warnings, or wants to see recent PyXLL activity.
- `pyxll_get_version` — before using any PyXLL feature that specifies a minimum version
  in the documentation, call this to confirm the feature is available.
- `pyxll_get_python_version` — when you need to confirm Python version compatibility
  before using language features or library APIs that require a specific version.
- `pyxll_get_config_path` — when you need to read or edit pyxll.cfg; call this first
  to find the file rather than guessing its location.
- `pyxll_get_config` — to inspect the parsed pyxll.cfg contents (includes variable
  substitutions and merged external configs); prefer this over reading the raw file.
- `pyxll_read_range(ref, sheet?)` — when the user asks to read or inspect cell values.
- `pyxll_read_formulas(ref, sheet?)` — when the user asks to see formulas in cells.
- `pyxll_write_range(ref, values, sheet?)` — when the user asks to write data or
  formulas into Excel cells.

## Writing Excel Functions

All functions belong in `pyxll_claude_functions.py`.  Decorate them with `@xl_func`
so PyXLL can find them:

```python
from pyxll import xl_func

@xl_func
def add_numbers(x: float, y: float) -> float:
    """Add two numbers — this docstring appears in Excel's function wizard."""
    return x + y
```

After writing or editing functions, call the `pyxll_reload_module` MCP tool to
register them with Excel immediately.  If it returns an error, fix the
error in `pyxll_claude_functions.py` and call `pyxll_reload_module` again.
Repeat until it returns `OK`.  Do not ask the user to click **Load Functions**.

## Workspace context

- User Excel functions: `pyxll_claude_functions.py` (in this workspace root)
- Use the `pyxll_reload_module` MCP tool to reload after edits.
- The **Load Functions** button in the task pane is a manual fallback only.

## Workspace layout

```
pyxll_claude_functions.py  ← your Excel functions (edit this file)
CLAUDE.md                  ← this file
.claude/
  skills/
    fetch-pyxll-docs/
      SKILL.md
```

## Rules

- After writing or editing functions in `pyxll_claude_functions.py`, call the
  `pyxll_reload_module` MCP tool. If it returns an error, fix it and call `pyxll_reload_module`
  again. Repeat until it returns `OK`. Do not ask the user to click **Load Functions**.
- Before using any PyXLL feature that the documentation marks with a minimum version
  requirement, call `pyxll_get_version` and confirm the installed version satisfies
  that requirement.  If it does not, use an alternative approach that works with the
  installed version, and tell the user what version would be needed for the preferred
  approach.
- Before calling `pyxll_reload_addin`, always call `pyxll_is_terminal_open` first.
  If it returns `True`, tell the user the terminal will close and ask them to confirm
  before proceeding — do NOT call `pyxll_reload_addin` until the user says yes.
  If it returns `False`, proceed without asking. Always prefer `pyxll_reload_module`
  for everyday edits — only escalate to `pyxll_reload_addin` when
  `pyxll_reload_module` is genuinely insufficient.
