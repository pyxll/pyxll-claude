import logging
import os
import traceback
from pathlib import Path

from pyxll import get_config

from . import register

_log = logging.getLogger(__name__)


def _get_log_path() -> Path:
    try:
        cfg = get_config()
        directory = os.path.expandvars(cfg.get("LOG", "path").strip())
        filename = os.path.expandvars(cfg.get("LOG", "file").strip())
        return Path(directory) / filename
    except Exception:
        _log.warning("could not read log path from pyxll config", exc_info=True)
        return Path(os.environ.get("APPDATA", "")) / "PyXLL" / "pyxll.log"


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    lines = int(args.get("lines", 50))
    try:
        log_path = _get_log_path()
        if not log_path.exists():
            return f"Log file not found: {log_path}", False
        with open(log_path, encoding="utf-8", errors="replace") as f:
            f.seek(0, 2)
            size = f.tell()
            chunk = min(size, lines * 200)
            f.seek(max(0, size - chunk))
            content = f.read()
        all_lines = content.splitlines()
        return "\n".join(all_lines[-lines:]), False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


register(
    "pyxll_get_log",
    (
        "Return the last N lines from the PyXLL log file. "
        "Use this whenever the user asks to see the PyXLL log, check for errors or "
        "warnings, or investigate recent PyXLL activity. Also useful after a reload "
        "to catch errors that appear in the log but not in the reload traceback."
    ),
    {
        "type": "object",
        "properties": {
            "lines": {
                "type": "integer",
                "description": "Number of lines to return from the end of the log. Defaults to 50.",
            }
        },
        "required": [],
    },
    _impl,
)
