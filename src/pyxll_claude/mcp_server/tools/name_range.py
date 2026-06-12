import traceback

from pyxll import xl_app

from . import register
from ._utils import run_on_main_thread


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    name = args.get("name")
    ref = args.get("ref")
    sheet = args.get("sheet")
    try:
        def _name():
            xl = xl_app()
            wb = xl.ActiveWorkbook
            if wb is None:
                raise ValueError("No active workbook")
            sheet_name = sheet or xl.ActiveSheet.Name
            refers_to = f"='{sheet_name}'!{ref}"
            try:
                wb.Names(name).RefersTo = refers_to
            except Exception:
                wb.Names.Add(Name=name, RefersTo=refers_to)

        run_on_main_thread(_name)
        return "OK", False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


register(
    "pyxll_name_range",
    (
        "Create or update a workbook-level named range. If a name already exists it is "
        "updated to point to the new reference; otherwise a new name is created. "
        "The range can then be referenced by name in formulas (e.g. =SUM(MyData))."
    ),
    {
        "type": "object",
        "properties": {
            "name": {
                "type": "string",
                "description": "The defined name to create or update, e.g. 'SalesData'.",
            },
            "ref": {
                "type": "string",
                "description": "A1 range reference, e.g. 'A1:D10'.",
            },
            "sheet": {
                "type": "string",
                "description": "Worksheet name the ref lives on. Omit to use the active sheet.",
            },
        },
        "required": ["name", "ref"],
    },
    _impl,
)
