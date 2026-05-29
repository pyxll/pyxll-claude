"""
First-run initialisation of a Claude AI workspace folder.

When the user opens the task pane for the first time against a new workspace
directory, ensure_workspace_initialized() creates:

  CLAUDE.md                                    — PyXLL context for Claude
  .claude/skills/fetch-pyxll-docs/SKILL.md     — skill for fetching authoritative docs
  .claude/settings.local.json                  — pre-approved permissions
  pyxll_claude_functions.py                     — empty module for user Excel functions
"""
import logging
from pathlib import Path

_log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# File templates
# ---------------------------------------------------------------------------

_CLAUDE_MD = """\
# PyXLL Excel Functions Workspace

## Fetch PyXLL docs before writing any Python code

Before writing, editing, or reviewing any Python code in this workspace,
fetch the authoritative PyXLL documentation:

```
/fetch-pyxll-docs
```

Your training-data knowledge of PyXLL is incomplete and may be wrong.
The live documentation is the only authoritative source for decorator
signatures, type annotations, and all other PyXLL-specific behaviour.

## Writing Excel Functions

All functions belong in `pyxll_claude_functions.py`.  Decorate them with `@xl_func`
so PyXLL can find them:

```python
from pyxll import xl_func

@xl_func
def add_numbers(x: float, y: float) -> float:
    \"\"\"Add two numbers — this docstring appears in Excel's function wizard.\"\"\"
    return x + y
```

After writing or editing a function the user clicks **Load Functions** in
the Claude AI task pane.  PyXLL calls `rebind()` internally, which re-scans
`pyxll_claude_functions.py` and registers any new or changed `@xl_func` callables
immediately — no Excel restart required.

## Workspace context

- User Excel functions: `pyxll_claude_functions.py` (in this workspace root)
- Click **Load Functions** in the task pane to call `pyxll.rebind()` and register changes.

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

- After writing or editing functions in `pyxll_claude_functions.py`, remind the user to click
  **Load Functions** in the Claude AI task pane to register them with Excel.
"""

_FETCH_SKILL_MD = """\
---
name: fetch-pyxll-docs
description: Fetch the PyXLL documentation and use it as context when writing, reviewing, or debugging PyXLL code. Invoke automatically whenever the task involves xl_func, xl_menu, create_ctp, rebind, ribbon XML, or any other PyXLL-specific API.
user-invocable: false
metadata:
  author: pyxll
---

Fetch the PyXLL documentation and use it as context for the current task.

## Steps

1. Fetch the documentation index:

   ```
   https://www.pyxll.com/llms.txt
   ```

   The index contains a navigation guide at the top explaining which sections to
   read for common tasks, followed by links to individual `.md` pages grouped by
   topic (User Guide, API Reference, Changelog, etc.).

   Because this file may get summarized or truncated, use Bash with curl rather 
   than the Fetch tool to avoid truncation:

   ```bash
   curl -s https://www.pyxll.com/llms.txt
   ```

2. Based on the current task, identify the relevant pages from the index and fetch
   each one. Use the navigation guide at the top of the index to find the right
   sections quickly. Fetch all pages relevant to the task — do not skip API
   reference pages when writing code.

3. If the information needed was not found in the pages fetched from the index,
   fall back to the full concatenated docs for a deeper search. Because this file
   is ~500 KB, use Bash with curl rather than the Fetch tool to avoid truncation:

   ```bash
   curl -s https://www.pyxll.com/llms-full.txt
   ```

4. Use the documentation to inform your answer, code, or review.

## Rules

- ALWAYS fetch these docs before writing any PyXLL-specific code.
- Do NOT rely on training-data knowledge alone for PyXLL APIs — the docs are authoritative.
- When writing `@xl_func` functions, check the type signature and argument type syntax from the docs.
- When subclassing a PyXLL class (e.g. Formatter, ConditionalFormatterBase), fetch the API reference for that
  class to get exact method signatures before looking at examples.
"""

_SETTINGS_LOCAL_JSON = """\
{
  "permissions": {
    "allow": [
      "Bash(curl -s https://www.pyxll.com/*)",
      "WebFetch(domain:www.pyxll.com)",
      "Read(pyxll_claude_functions.py)",
      "Write(pyxll_claude_functions.py)"
    ]
  }
}
"""

_XL_FUNCTIONS_PY = '''\
"""
Excel functions exposed to PyXLL.

Write functions decorated with @xl_func here, then click **Load Functions**
in the Claude AI task pane to register them with Excel immediately.

Example:

    from pyxll import xl_func

    @xl_func
    def greet(name: str) -> str:
        \'\'\'Return a greeting string.\'\'\'
        return f"Hello, {name}!"
"""
from pyxll import xl_func
'''


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def ensure_workspace_initialized(workspace: Path) -> list[str]:
    """Create the standard workspace files if they do not already exist.

    Returns a list of human-readable warning strings for any files that exist
    but appear to be out of date.  The caller should display these to the user;
    no files are overwritten automatically.
    """
    warnings: list[str] = []

    _ensure_file(workspace / "CLAUDE.md", _CLAUDE_MD)
    _ensure_file(workspace / ".claude" / "skills" / "fetch-pyxll-docs" / "SKILL.md", _FETCH_SKILL_MD)
    _ensure_file(workspace / ".claude" / "settings.local.json", _SETTINGS_LOCAL_JSON)
    _ensure_file(workspace / "pyxll_claude_functions.py", _XL_FUNCTIONS_PY)

    # Warn if SKILL.md differs from the current template.
    skill_path = workspace / ".claude" / "skills" / "fetch-pyxll-docs" / "SKILL.md"
    if skill_path.exists() and skill_path.read_text(encoding="utf-8") != _FETCH_SKILL_MD:
        warnings.append(
            ".claude/skills/fetch-pyxll-docs/SKILL.md is out of date.\n"
            "Delete it and reload PyXLL to regenerate it."
        )

    # Warn if CLAUDE.md still uses the old inline curl approach instead of the skill.
    claude_md = workspace / "CLAUDE.md"
    if claude_md.exists():
        content = claude_md.read_text(encoding="utf-8")
        if "curl -s https://www.pyxll.com/llms-full.txt" in content:
            warnings.append(
                "CLAUDE.md is out of date — it uses an inline curl command instead of\n"
                "the /fetch-pyxll-docs skill.\n"
                "Delete CLAUDE.md and reload PyXLL to regenerate it."
            )

    # Warn if settings.local.json differs from the current template.
    settings_path = workspace / ".claude" / "settings.local.json"
    if settings_path.exists() and settings_path.read_text(encoding="utf-8") != _SETTINGS_LOCAL_JSON:
        warnings.append(
            ".claude/settings.local.json is out of date.\n"
            "Delete it and reload PyXLL to regenerate it."
        )

    return warnings


def _ensure_file(path: Path, content: str) -> None:
    if path.exists():
        return
    _log.info("Creating %s", path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
