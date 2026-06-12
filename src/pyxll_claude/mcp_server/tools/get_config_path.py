import os
import traceback
from pathlib import Path

import pyxll

from . import register


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    try:
        env = os.environ.get("PYXLL_CONFIG_FILE", "").strip()
        if env:
            return env, False
        xll = Path(pyxll.__file__)
        return str(xll.with_suffix(".cfg")), False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


register(
    "pyxll_get_config_path",
    (
        "Return the full path (including filename) of the pyxll.cfg configuration file. "
        "Checks the PYXLL_CONFIG_FILE environment variable first; if unset, derives the "
        "path from pyxll.__file__ by replacing its extension with .cfg (e.g. pyxll.xll "
        "→ pyxll.cfg). Use this before reading or editing pyxll.cfg."
    ),
    {"type": "object", "properties": {}, "required": []},
    _impl,
)
