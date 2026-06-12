import traceback

from pyxll import xl_app

from . import register
from ._utils import get_sheet, run_on_main_thread


def _merge_impl(args: dict, ctx: dict) -> tuple[str, bool]:
    ref = args.get("ref")
    sheet = args.get("sheet")
    try:
        def _merge():
            xl = xl_app()
            get_sheet(xl, sheet).Range(ref).Merge()

        run_on_main_thread(_merge)
        return "OK", False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


def _unmerge_impl(args: dict, ctx: dict) -> tuple[str, bool]:
    ref = args.get("ref")
    sheet = args.get("sheet")
    try:
        def _unmerge():
            xl = xl_app()
            get_sheet(xl, sheet).Range(ref).UnMerge()

        run_on_main_thread(_unmerge)
        return "OK", False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


_SCHEMA = {
    "type": "object",
    "properties": {
        "ref": {
            "type": "string",
            "description": "A1 range reference, e.g. 'A1:C1'.",
        },
        "sheet": {
            "type": "string",
            "description": "Worksheet name. Omit to use the active sheet.",
        },
    },
    "required": ["ref"],
}

register(
    "pyxll_merge_cells",
    "Merge the cells of a range into a single cell.",
    _SCHEMA,
    _merge_impl,
)

register(
    "pyxll_unmerge_cells",
    "Unmerge previously merged cells, splitting them back into individual cells.",
    _SCHEMA,
    _unmerge_impl,
)
