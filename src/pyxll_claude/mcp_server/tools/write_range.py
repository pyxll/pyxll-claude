import threading
import traceback
from typing import Any

import pyxll
from pyxll import xl_app

from . import register


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    ref = args.get("ref")
    values = args.get("values", [])
    sheet = args.get("sheet")
    result: list[str | None] = [None]
    done = threading.Event()

    def _do_write():
        try:
            xl = xl_app()
            ws = xl.Sheets(sheet) if sheet else xl.ActiveSheet
            top_left = ws.Range(ref)
            rows = len(values)
            cols = max((len(r) for r in values), default=0)
            rng = top_left.GetResize(rows, cols)
            # Pad each row to uniform width as a list of lists.
            # win32com requires a true 2-D list-of-lists to produce the
            # correct SAFEARRAY shape; tuple-of-tuples collapses to 1-D
            # for single-column ranges and writes to the wrong cells.
            padded = [list(row) + [None] * (cols - len(row)) for row in values]
            has_formula = any(
                isinstance(v, str) and v.startswith("=")
                for row in padded
                for v in row
            )
            if has_formula:
                for r_idx, row in enumerate(padded):
                    for c_idx, v in enumerate(row):
                        cell = top_left.GetOffset(r_idx, c_idx)
                        if isinstance(v, str) and v.startswith("="):
                            cell.Formula2 = v
                        else:
                            cell.Value2 = v
            else:
                rng.Value2 = padded
            result[0] = "OK"
        except Exception:
            result[0] = "ERROR: " + traceback.format_exc()
        finally:
            done.set()

    pyxll.schedule_call(_do_write)
    done.wait(timeout=15)

    if result[0] is None:
        return "ERROR: write timed out", True
    text = result[0]
    return text, text != "OK"


register(
    "pyxll_write_range",
    (
        "Write values to an Excel worksheet range. "
        "Strings starting with '=' are written as formulas."
    ),
    {
        "type": "object",
        "properties": {
            "ref": {
                "type": "string",
                "description": "Top-left cell of the destination range, e.g. 'B3'.",
            },
            "values": {
                "type": "array",
                "items": {"type": "array"},
                "description": "2-D array of values (list of rows).",
            },
            "sheet": {
                "type": "string",
                "description": "Worksheet name. Omit to use the active sheet.",
            },
        },
        "required": ["ref", "values"],
    },
    _impl,
)
