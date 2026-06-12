import traceback

from pyxll import xl_app

from . import register
from ._utils import get_sheet, run_on_main_thread


def _autofit_impl(args: dict, ctx: dict) -> tuple[str, bool]:
    column = args.get("column")
    sheet = args.get("sheet")
    try:
        def _autofit():
            xl = xl_app()
            col_range = column if ":" in column else f"{column}:{column}"
            get_sheet(xl, sheet).Range(col_range).Columns.AutoFit()

        run_on_main_thread(_autofit)
        return "OK", False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


def _set_width_impl(args: dict, ctx: dict) -> tuple[str, bool]:
    column = args.get("column")
    width = args.get("width")
    sheet = args.get("sheet")
    try:
        def _set():
            xl = xl_app()
            col_range = column if ":" in column else f"{column}:{column}"
            get_sheet(xl, sheet).Range(col_range).ColumnWidth = width

        run_on_main_thread(_set)
        return "OK", False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


def _set_height_impl(args: dict, ctx: dict) -> tuple[str, bool]:
    row = args.get("row")
    height = args.get("height")
    sheet = args.get("sheet")
    try:
        def _set():
            xl = xl_app()
            get_sheet(xl, sheet).Rows(row).RowHeight = height

        run_on_main_thread(_set)
        return "OK", False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


_SHEET_PROP = {
    "type": "string",
    "description": "Worksheet name. Omit to use the active sheet.",
}

register(
    "pyxll_auto_fit_column",
    (
        "Auto-fit one or more columns to their content width. "
        "Pass a single column letter ('A') or a column range ('A:C')."
    ),
    {
        "type": "object",
        "properties": {
            "column": {
                "type": "string",
                "description": "Column letter or range, e.g. 'A' or 'A:C'.",
            },
            "sheet": _SHEET_PROP,
        },
        "required": ["column"],
    },
    _autofit_impl,
)

register(
    "pyxll_set_column_width",
    (
        "Set the width of a column or column range in character units. "
        "Pass a single letter ('B') or a range ('B:D') and the desired width."
    ),
    {
        "type": "object",
        "properties": {
            "column": {
                "type": "string",
                "description": "Column letter or range, e.g. 'B' or 'B:D'.",
            },
            "width": {
                "type": "number",
                "description": "Column width in character units.",
            },
            "sheet": _SHEET_PROP,
        },
        "required": ["column", "width"],
    },
    _set_width_impl,
)

register(
    "pyxll_set_row_height",
    (
        "Set the height of a row in points."
    ),
    {
        "type": "object",
        "properties": {
            "row": {
                "type": "integer",
                "description": "1-indexed row number.",
            },
            "height": {
                "type": "number",
                "description": "Row height in points.",
            },
            "sheet": _SHEET_PROP,
        },
        "required": ["row", "height"],
    },
    _set_height_impl,
)
