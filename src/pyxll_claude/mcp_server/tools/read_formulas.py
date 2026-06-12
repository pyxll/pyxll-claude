import json
import threading
import traceback
from typing import Any

import pyxll
from pyxll import xl_app

from . import register


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    ref = args.get("ref")
    sheet = args.get("sheet")
    result: list[Any] = [None]
    done = threading.Event()

    def _do_read():
        try:
            xl = xl_app()
            ws = xl.Sheets(sheet) if sheet else xl.ActiveSheet
            v = ws.Range(ref).Formula2
            if not isinstance(v, tuple):
                v = ((v,),)
            result[0] = json.dumps([list(row) for row in v])
        except Exception:
            result[0] = "ERROR: " + traceback.format_exc()
        finally:
            done.set()

    pyxll.schedule_call(_do_read)
    done.wait(timeout=15)

    if result[0] is None:
        return "ERROR: read timed out", True
    text = result[0]
    return text, text.startswith("ERROR:")


register(
    "pyxll_read_formulas",
    (
        "Read cell formulas from an Excel worksheet using the Formula2 property. "
        "Returns a JSON-encoded 2-D list (list of rows). "
        "Formula cells return the formula string (e.g. '=SUM(A1:A10)'); "
        "non-formula cells return the raw value. "
        "Single cells return [[value]]."
    ),
    {
        "type": "object",
        "properties": {
            "ref": {
                "type": "string",
                "description": "A1 range reference, e.g. 'A1' or 'B2:D5'.",
            },
            "sheet": {
                "type": "string",
                "description": "Worksheet name, e.g. 'Sheet1'. Omit to use the active sheet.",
            },
        },
        "required": ["ref"],
    },
    _impl,
)
