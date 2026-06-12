import traceback

from pyxll import xl_app

from . import register
from ._utils import run_on_main_thread


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    name = args.get("name")
    after = args.get("after")
    try:
        def _add():
            xl = xl_app()
            wb = xl.ActiveWorkbook
            if wb is None:
                raise ValueError("No active workbook")
            anchor = wb.Sheets(after) if after else wb.Sheets(wb.Sheets.Count)
            new_sheet = wb.Sheets.Add(After=anchor)
            if name:
                new_sheet.Name = name
            return new_sheet.Name

        return run_on_main_thread(_add), False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


register(
    "pyxll_add_sheet",
    (
        "Add a new worksheet to the active workbook and return its name. "
        "The sheet is inserted after the last sheet by default, or after 'after' if given. "
        "Returns the actual sheet name (useful when Excel adjusts it for uniqueness)."
    ),
    {
        "type": "object",
        "properties": {
            "name": {
                "type": "string",
                "description": "Name for the new sheet. Excel assigns a default name if omitted.",
            },
            "after": {
                "type": "string",
                "description": "Name of the sheet after which to insert. Defaults to the last sheet.",
            },
        },
        "required": [],
    },
    _impl,
)
