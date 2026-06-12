import traceback

import pyxll

from . import register


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    try:
        return pyxll.__version__, False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


register(
    "pyxll_get_version",
    (
        "Return the installed PyXLL version string (e.g. '5.4.1'). "
        "Call this before using any PyXLL feature that specifies a minimum version "
        "in the documentation, to confirm the feature is available."
    ),
    {"type": "object", "properties": {}, "required": []},
    _impl,
)
