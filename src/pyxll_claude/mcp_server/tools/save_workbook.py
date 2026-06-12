import traceback

from pyxll import xl_app

from . import register
from ._utils import run_on_main_thread


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    path = args.get("path")
    try:
        def _save():
            xl = xl_app()
            wb = xl.ActiveWorkbook
            if wb is None:
                raise ValueError("No active workbook")
            if path:
                wb.SaveAs(path)
            elif wb.Path:
                wb.Save()
            else:
                raise ValueError("Workbook has never been saved; provide a path.")
            return wb.FullName

        return run_on_main_thread(_save, timeout=30.0), False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


register(
    "pyxll_save_workbook",
    (
        "Save the active workbook and return its full file path. "
        "If 'path' is given the workbook is saved to that location (SaveAs); "
        "otherwise it is saved in place. Raises an error if the workbook has never "
        "been saved and no path is provided."
    ),
    {
        "type": "object",
        "properties": {
            "path": {
                "type": "string",
                "description": "Full file path for SaveAs. Omit to save in place.",
            },
        },
        "required": [],
    },
    _impl,
)
