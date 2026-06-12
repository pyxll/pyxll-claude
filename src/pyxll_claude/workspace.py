"""
First-run initialisation of a Claude AI workspace folder.

When the user opens the task pane for the first time against a new workspace
directory, ensure_workspace_initialized() creates:

  CLAUDE.md                                                    — PyXLL context for Claude
  .claude/settings.local.json                                  — pre-approved permissions
  pyxll_claude_functions.py                                    — empty module for user Excel functions

Skills are always synced to the current package version:

  .claude/skills/<skill>/...                                   — mirrored from resources/workspace/skills/
"""
import logging
import stat
from importlib.resources import files
from importlib.resources.abc import Traversable
from pathlib import Path

_log = logging.getLogger(__name__)

_WORKSPACE_RES = "pyxll_claude.resources"

# Files copied once on first run (never overwritten).
_ONCE_FILES: list[tuple[str, str]] = [
    ("CLAUDE.md",                  "CLAUDE.md"),
    ("pyxll_claude_functions.py",  "pyxll_claude_functions.py"),
    ("claude.settings.local.json", ".claude/settings.local.json"),
]

_EXECUTABLE_SUFFIXES = {".sh"}


def ensure_workspace_initialized(workspace: Path) -> list[str]:
    """Create and sync workspace files.

    Files in _ONCE_FILES are created only if absent.
    Everything under resources/workspace/skills/ is always synced to
    workspace/.claude/skills/ so skills stay up to date automatically.

    Returns a list of human-readable warning strings for the caller to display.
    """
    warnings: list[str] = []

    res_root = files(_WORKSPACE_RES).joinpath("workspace")

    for src_name, dest_rel in _ONCE_FILES:
        content = res_root.joinpath(src_name).read_text(encoding="utf-8")
        _ensure_file(workspace / dest_rel, content)

    _sync_dir(
        res_root.joinpath("skills"),
        workspace / ".claude" / "skills",
    )

    return warnings


def _sync_dir(src: Traversable, dest: Path) -> None:
    """Recursively sync src (package Traversable) into dest, overwriting changed files."""
    dest.mkdir(parents=True, exist_ok=True)
    for item in src.iterdir():
        dest_item = dest / item.name
        if item.is_dir():
            _sync_dir(item, dest_item)
        else:
            content = item.read_text(encoding="utf-8")
            if not dest_item.exists() or dest_item.read_text(encoding="utf-8") != content:
                _log.info("Writing %s", dest_item)
                dest_item.parent.mkdir(parents=True, exist_ok=True)
                dest_item.write_text(content, encoding="utf-8")
            if Path(item.name).suffix in _EXECUTABLE_SUFFIXES:
                dest_item.chmod(dest_item.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)


def _ensure_file(path: Path, content: str) -> None:
    if path.exists():
        return
    _log.info("Creating %s", path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
