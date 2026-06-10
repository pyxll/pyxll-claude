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
  Call this after every edit to `pyxll_claude_functions.py`.
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
    \"\"\"Add two numbers — this docstring appears in Excel's function wizard.\"\"\"
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

1. Fetch the documentation index to find relevant pages:

   ```bash
   curl -s https://www.pyxll.com/llms.txt
   ```

   Read the **entire** output without truncating, piping to `head`, or
   summarising. The index contains a navigation guide at the top mapping
   common tasks to sections, followed by page titles, descriptions, and URLs
   grouped by topic. Truncating it will cause you to miss relevant pages.

2. Fetch the individual pages relevant to the task directly by their URL:

   ```bash
   curl -s <page-url>
   ```

3. If the index does not surface what you need, use the search script to find
   pages by keyword. The script caches the full docs locally (refreshed every
   24 h) and returns only matching page URLs — avoiding loading 500 KB into context.

   The script is in the `scripts/` folder next to this file. Find this file's
   location and run:

   ```bash
   /path/to/this/skills/fetch-pyxll-docs/scripts/search-pyxll-docs.sh <keyword> [keyword2 ...]
   ```

   Then fetch the returned page URLs individually using curl.

4. Use the documentation to inform your answer, code, or review.

## Rules

- ALWAYS fetch these docs before writing, modifying, or troubleshooting any
  PyXLL-specific code or behaviour. Before suggesting a manual workaround for a
  PyXLL problem, check whether PyXLL already has a built-in solution (decorator
  parameter, config key, or feature).
- Do NOT rely on training-data knowledge alone for PyXLL APIs — the docs are authoritative.
- When writing `@xl_func` functions, check the type signature and argument type
  syntax from the docs.
- Before writing any code that calls the Excel COM API (Range, Worksheet, Workbook, etc.),
  fetch https://www.pyxll.com/docs/userguide/vba.md and read it in full. It documents
  critical differences between VBA and Python — including how COM properties that take
  arguments must be called as `Get<PropertyName>(args)` in Python rather than
  `Property(args)` as in VBA.
- Before using any PyXLL class, function, decorator, or configuration setting
  (including pyxll.cfg section names and their keys), fetch the relevant documentation
  and use only what is explicitly documented. Never infer behaviour, key names, or
  parameter names from conventions or assumptions — if it is not in the docs, do not
  use it.
"""

_SEARCH_SCRIPT_SH = """\
#!/usr/bin/env bash
# search-pyxll-docs.sh -- Search PyXLL full docs, return matching page URLs.
#
# Downloads and caches the full docs locally (refreshed every 24 h) so only
# the relevant pages need to be loaded into the AI context.
#
# Usage: ./search-pyxll-docs.sh <keyword> [keyword2 ...] [-f|--fresh]

set -euo pipefail

FULL_URL="https://www.pyxll.com/llms-full.txt"
CACHE_DIR="${XDG_CACHE_HOME:-$HOME/.cache}/pyxll-docs"
CACHE_FILE="$CACHE_DIR/llms-full.txt"
CACHE_MAX_AGE=86400

usage() { echo "Usage: $0 [-f|--fresh] <keyword> [keyword2 ...]"; exit 1; }

FRESH=false
KEYWORDS=()
while [[ $# -gt 0 ]]; do
    case $1 in
        -f|--fresh) FRESH=true; shift ;;
        -h|--help)  usage ;;
        -*)         echo "Unknown option: $1"; usage ;;
        *)          KEYWORDS+=("$1"); shift ;;
    esac
done
[[ ${#KEYWORDS[@]} -eq 0 ]] && { echo "Error: at least one keyword required"; usage; }

mkdir -p "$CACHE_DIR"

need_download() {
    [[ "$FRESH" == "true" ]] && return 0
    [[ ! -f "$CACHE_FILE" ]] && return 0
    local age
    if [[ "$(uname)" == "Darwin" ]]; then
        age=$(( $(date +%s) - $(stat -f %m "$CACHE_FILE") ))
    else
        age=$(( $(date +%s) - $(stat -c %Y "$CACHE_FILE") ))
    fi
    [[ $age -gt $CACHE_MAX_AGE ]]
}

if need_download; then
    echo "Downloading PyXLL docs (cached for 24 h)..." >&2
    curl -sL "$FULL_URL" -o "$CACHE_FILE"
fi

kw_pattern=""
for kw in "${KEYWORDS[@]}"; do
    kw_pattern="${kw_pattern}|$(printf '%s' "$kw" | tr '[:upper:]' '[:lower:]')"
done
kw_pattern="${kw_pattern:1}"

echo "Searching for: ${KEYWORDS[*]}" >&2

# PyXLL llms-full.txt page format:
#   ---                                    <- page boundary
#   ## Title                               <- page header (## not ###)
#   [path.md](https://www.pyxll.com/...)  <- URL line immediately after title
awk -v kw="$kw_pattern" '
BEGIN { url=""; n=split(kw, K, "|"); state=0 }
/^---$/                                { state=1; next }
state==1 && /^## /                     { state=2; next }
state==2 && index($0, "www.pyxll.com") {
    if (split($0, a, "(") >= 2) {
        cp = index(a[2], ")")
        if (cp > 0) url = substr(a[2], 1, cp-1)
    }
    state=0; next
}
url != "" {
    line = tolower($0)
    for (i=1; i<=n; i++) {
        if (index(line, K[i]) > 0 && !seen[url]) {
            seen[url]=1; print url; break
        }
    }
}
' "$CACHE_FILE"
"""

_SETTINGS_LOCAL_JSON = """\
{
  "permissions": {
    "allow": [
      "Bash(*/search-pyxll-docs.sh*)",
      "Bash(curl -s https://www.pyxll.com/*)",
      "Read(pyxll_claude_functions.py)",
      "Write(pyxll_claude_functions.py)",
      "mcp__pyxll__pyxll_reload_module",
      "mcp__pyxll__pyxll_get_log",
      "mcp__pyxll__pyxll_get_version",
      "mcp__pyxll__pyxll_get_python_version",
      "mcp__pyxll__pyxll_get_config_path",
      "mcp__pyxll__pyxll_get_config",
      "mcp__pyxll__pyxll_get_selection",
      "mcp__pyxll__pyxll_read_range",
      "mcp__pyxll__pyxll_read_formulas",
      "mcp__pyxll__pyxll_write_range"
    ]
  },
  "enabledMcpjsonServers": [
    "pyxll"
  ]
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

    skill_dir = workspace / ".claude" / "skills" / "fetch-pyxll-docs"
    skill_path = skill_dir / "SKILL.md"
    script_path = skill_dir / "scripts" / "search-pyxll-docs.sh"
    settings_path = workspace / ".claude" / "settings.local.json"

    _ensure_file(workspace / "CLAUDE.md", _CLAUDE_MD)
    _ensure_file(skill_path, _FETCH_SKILL_MD)
    _ensure_file(script_path, _SEARCH_SCRIPT_SH)
    _ensure_file(settings_path, _SETTINGS_LOCAL_JSON)
    _ensure_file(workspace / "pyxll_claude_functions.py", _XL_FUNCTIONS_PY)

    # Ensure the search script is executable.
    if script_path.exists():
        script_path.chmod(script_path.stat().st_mode | 0o111)

    # Warn if any generated file differs from the current template.
    if _file_differs(skill_path, _FETCH_SKILL_MD):
        warnings.append(
            ".claude/skills/fetch-pyxll-docs/SKILL.md is out of date.\n"
            "Delete it and restart Claude to regenerate it."
        )

    if _file_differs(script_path, _SEARCH_SCRIPT_SH):
        warnings.append(
            ".claude/skills/fetch-pyxll-docs/scripts/search-pyxll-docs.sh"
            " is out of date.\n"
            "Delete it and restart Claude to regenerate it."
        )

    if _file_differs(workspace / "CLAUDE.md", _CLAUDE_MD):
        warnings.append(
            "CLAUDE.md is out of date.\n"
            "Delete it and restart Claude to regenerate it."
        )

    if _file_differs(settings_path, _SETTINGS_LOCAL_JSON):
        warnings.append(
            ".claude/settings.local.json is out of date.\n"
            "Delete it and restart Claude to regenerate it."
        )

    return warnings


def _file_differs(path: Path, content: str) -> bool:
    """Return True if path exists but its content differs from content."""
    return path.exists() and path.read_text(encoding="utf-8") != content


def _ensure_file(path: Path, content: str) -> None:
    if path.exists():
        return
    _log.info("Creating %s", path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
