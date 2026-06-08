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

1. Fetch the documentation index to find relevant pages:

   ```bash
   curl -s https://www.pyxll.com/llms.txt
   ```

   The index contains a navigation guide at the top mapping common tasks to
   sections, followed by page titles, descriptions, and URLs grouped by topic.

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

- ALWAYS fetch these docs before writing any PyXLL-specific code.
- Do NOT rely on training-data knowledge alone for PyXLL APIs — the docs are authoritative.
- When writing `@xl_func` functions, check the type signature and argument type syntax from the docs.
- Before using any PyXLL class, function, or decorator, fetch its API reference and use
  only what is explicitly documented. Never infer behaviour from conventions or assumptions —
  if it is not in the docs, do not use it.
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
    _ensure_file(workspace / ".claude" / "skills" / "fetch-pyxll-docs" / "scripts" / "search-pyxll-docs.sh", _SEARCH_SCRIPT_SH)
    _ensure_file(workspace / ".claude" / "settings.local.json", _SETTINGS_LOCAL_JSON)
    _ensure_file(workspace / "pyxll_claude_functions.py", _XL_FUNCTIONS_PY)

    # Ensure the search script is executable.
    script_path = workspace / ".claude" / "skills" / "fetch-pyxll-docs" / "scripts" / "search-pyxll-docs.sh"
    if script_path.exists():
        script_path.chmod(script_path.stat().st_mode | 0o111)

    # Warn if SKILL.md differs from the current template.
    skill_path = workspace / ".claude" / "skills" / "fetch-pyxll-docs" / "SKILL.md"
    if skill_path.exists() and skill_path.read_text(encoding="utf-8") != _FETCH_SKILL_MD:
        warnings.append(
            ".claude/skills/fetch-pyxll-docs/SKILL.md is out of date.\n"
            "Delete it and reload PyXLL to regenerate it."
        )

    # Warn if the search script differs from the current template.
    if script_path.exists() and script_path.read_text(encoding="utf-8") != _SEARCH_SCRIPT_SH:
        warnings.append(
            ".claude/skills/fetch-pyxll-docs/scripts/search-pyxll-docs.sh is out of date.\n"
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
