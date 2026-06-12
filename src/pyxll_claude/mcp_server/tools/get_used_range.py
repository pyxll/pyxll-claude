import json
import traceback

from pyxll import xl_app

from . import register
from ._utils import get_sheet, run_on_main_thread


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    sheet = args.get("sheet")
    try:
        def _get():
            xl = xl_app()
            used = get_sheet(xl, sheet).UsedRange
            return {
                "address": used.Address.replace("$", ""),
                "row_count": used.Rows.Count,
                "column_count": used.Columns.Count,
                "start_row": used.Row,
                "start_column": used.Column,
            }

        return json.dumps(run_on_main_thread(_get)), False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


register(
    "pyxll_get_used_range",
    (
        "Return the bounds of the populated area of a worksheet as a JSON object with "
        "'address', 'row_count', 'column_count', 'start_row', and 'start_column' fields. "
        "Use this to discover the extent of data before iterating or reading a large range."
    ),
    {
        "type": "object",
        "properties": {
            "sheet": {
                "type": "string",
                "description": "Worksheet name. Omit to use the active sheet.",
            },
        },
        "required": [],
    },
    _impl,
)
