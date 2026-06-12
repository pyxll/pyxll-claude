import json
import threading
import traceback
from typing import Any

import pyxll
from pyxll import xl_app

from . import register


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    result: list[Any] = [None]
    done = threading.Event()

    def _do_get():
        try:
            xl = xl_app()
            sel = xl.Selection
            if sel is None:
                result[0] = "ERROR: no selection"
                return
            address = sel.Address.replace("$", "")
            sheet = xl.ActiveSheet.Name
            result[0] = json.dumps({"address": address, "sheet": sheet})
        except Exception:
            result[0] = "ERROR: " + traceback.format_exc()
        finally:
            done.set()

    pyxll.schedule_call(_do_get)
    done.wait(timeout=15)

    if result[0] is None:
        return "ERROR: timed out", True
    text = result[0]
    return text, text.startswith("ERROR:")


register(
    "pyxll_get_selection",
    (
        "Return the current Excel selection as a JSON object with 'address' and 'sheet' fields. "
        "Call this first whenever you need to know which cell(s) the user has selected — "
        "e.g. 'check the current cell', 'read the selected range', or any read/write operation "
        "relative to the selection. Pass the returned address to pyxll_read_range or "
        "pyxll_write_range. Works for single cells, ranges, and non-contiguous selections."
    ),
    {"type": "object", "properties": {}, "required": []},
    _impl,
)
