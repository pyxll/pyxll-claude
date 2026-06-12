import traceback

from pyxll import xl_app

from . import register
from ._utils import get_sheet, run_on_main_thread


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    ref = args.get("ref")
    clear_type = args.get("clear_type", "all")
    sheet = args.get("sheet")
    try:
        def _clear():
            xl = xl_app()
            rng = get_sheet(xl, sheet).Range(ref)
            if clear_type == "contents":
                rng.ClearContents()
            elif clear_type == "formats":
                rng.ClearFormats()
            else:
                rng.Clear()

        run_on_main_thread(_clear)
        return "OK", False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


register(
    "pyxll_clear_range",
    (
        "Clear contents, formats, or both from a range without deleting rows or columns. "
        "Use clear_type='contents' to erase values and formulas only, 'formats' to remove "
        "formatting only, or 'all' (default) to clear everything."
    ),
    {
        "type": "object",
        "properties": {
            "ref": {
                "type": "string",
                "description": "A1 range reference, e.g. 'A1:C10'.",
            },
            "clear_type": {
                "type": "string",
                "enum": ["all", "contents", "formats"],
                "description": "What to clear. Defaults to 'all'.",
            },
            "sheet": {
                "type": "string",
                "description": "Worksheet name. Omit to use the active sheet.",
            },
        },
        "required": ["ref"],
    },
    _impl,
)
