import traceback

from pyxll import xl_app

from . import register
from ._utils import get_sheet, run_on_main_thread


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    scope = args.get("scope", "workbook")
    sheet = args.get("sheet")
    ref = args.get("ref")
    try:
        def _calc():
            xl = xl_app()
            if scope == "workbook":
                # CalculateFull forces recalculation regardless of automatic/manual mode.
                xl.CalculateFull()
            elif scope == "sheet":
                ws = get_sheet(xl, sheet)
                ws.UsedRange.Dirty()
                ws.Calculate()
            elif scope == "range":
                if not ref:
                    raise ValueError("ref is required when scope is 'range'")
                rng = get_sheet(xl, sheet).Range(ref)
                rng.Dirty()
                rng.Calculate()
            else:
                raise ValueError(f"Invalid scope: {scope!r}")

        run_on_main_thread(_calc)
        return "OK", False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


register(
    "pyxll_calculate",
    (
        "Force a recalculation at workbook, sheet, or range scope. "
        "Use scope='workbook' (default) to recalculate everything, 'sheet' for one sheet, "
        "or 'range' (with ref) for a specific range. Useful when automatic calculation "
        "is disabled or after writing values that affect formulas."
    ),
    {
        "type": "object",
        "properties": {
            "scope": {
                "type": "string",
                "enum": ["workbook", "sheet", "range"],
                "description": "What to recalculate. Defaults to 'workbook'.",
            },
            "sheet": {
                "type": "string",
                "description": "Worksheet name. Required when scope is 'sheet' or 'range'.",
            },
            "ref": {
                "type": "string",
                "description": "A1 range reference. Required when scope is 'range'.",
            },
        },
        "required": [],
    },
    _impl,
)
