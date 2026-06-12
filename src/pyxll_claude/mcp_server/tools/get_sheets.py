import json
import traceback

from pyxll import xl_app

from . import register
from ._utils import run_on_main_thread


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    try:
        def _get():
            xl = xl_app()
            if xl.ActiveWorkbook is None:
                return []
            return [s.Name for s in xl.ActiveWorkbook.Sheets]

        return json.dumps(run_on_main_thread(_get)), False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


register(
    "pyxll_get_sheets",
    (
        "Return a JSON array of worksheet names in the active workbook. "
        "Call this to discover available sheets before reading from or writing to a "
        "specific sheet by name."
    ),
    {"type": "object", "properties": {}, "required": []},
    _impl,
)
