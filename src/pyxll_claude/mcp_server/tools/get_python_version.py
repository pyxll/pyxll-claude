import sys
import traceback

from . import register


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    try:
        return sys.version, False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


register(
    "pyxll_get_python_version",
    (
        "Return the Python version string (e.g. '3.11.4 (main, ...')). "
        "Call this when you need to confirm Python version compatibility "
        "before using language features or library APIs that require a specific version."
    ),
    {"type": "object", "properties": {}, "required": []},
    _impl,
)
