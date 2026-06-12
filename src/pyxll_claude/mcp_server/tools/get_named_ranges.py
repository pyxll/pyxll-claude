import json
import traceback

from pyxll import xl_app

from . import register
from ._utils import run_on_main_thread


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    try:
        def _get():
            xl = xl_app()
            wb = xl.ActiveWorkbook
            if wb is None:
                return []
            return [{"name": n.Name, "refers_to": n.RefersTo} for n in wb.Names]

        return json.dumps(run_on_main_thread(_get)), False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


register(
    "pyxll_get_named_ranges",
    (
        "Return a JSON array of all defined names in the active workbook, each with "
        "'name' and 'refers_to' fields. Use this to discover named ranges, named "
        "constants, and named formulas before referencing them by name."
    ),
    {"type": "object", "properties": {}, "required": []},
    _impl,
)
