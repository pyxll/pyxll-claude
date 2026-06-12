import io
import traceback

from pyxll import get_config

from . import register


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    try:
        buf = io.StringIO()
        get_config().write(buf)
        return buf.getvalue(), False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


register(
    "pyxll_get_config",
    (
        "Return the complete, parsed contents of pyxll.cfg as an INI-formatted string. "
        "Uses pyxll.get_config() so variable substitutions and any externally merged "
        "config files are already applied. Prefer this over reading the raw file when "
        "you need to inspect or reason about PyXLL configuration values."
    ),
    {"type": "object", "properties": {}, "required": []},
    _impl,
)
