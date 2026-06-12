import traceback

from pyxll import xl_app

from . import register
from ._utils import get_sheet, run_on_main_thread

_MAX_INSERT = 1000


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    row = args.get("row")
    count = int(args.get("count", 1))
    sheet = args.get("sheet")
    if not 1 <= count <= _MAX_INSERT:
        return f"ERROR: count must be between 1 and {_MAX_INSERT}, got {count}", True
    try:
        def _insert():
            xl = xl_app()
            ws = get_sheet(xl, sheet)
            for _ in range(count):
                ws.Rows(row).Insert()

        run_on_main_thread(_insert)
        return "OK", False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


register(
    "pyxll_insert_rows",
    (
        "Insert one or more blank rows above the given row number (1-indexed), "
        "shifting existing rows down. Use this to make room for new data without "
        "overwriting existing content."
    ),
    {
        "type": "object",
        "properties": {
            "row": {
                "type": "integer",
                "description": "1-indexed row number above which to insert.",
            },
            "count": {
                "type": "integer",
                "description": "Number of rows to insert. Defaults to 1.",
            },
            "sheet": {
                "type": "string",
                "description": "Worksheet name. Omit to use the active sheet.",
            },
        },
        "required": ["row"],
    },
    _impl,
)
