import importlib
import logging
import sys
import threading
import traceback

import pyxll
from pyxll import rebind

from . import register

_log = logging.getLogger(__name__)


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    module_name = args.get("module") or "pyxll_claude_functions"
    workspace = ctx["workspace"]
    result: list[str | None] = [None]
    done = threading.Event()
    ws_str = str(workspace)

    def _do_reload():
        try:
            if ws_str not in sys.path:
                sys.path.insert(0, ws_str)
            if module_name in sys.modules:
                importlib.reload(sys.modules[module_name])
            else:
                importlib.import_module(module_name)
            rebind()
            result[0] = "OK"
        except Exception:
            result[0] = traceback.format_exc()
        finally:
            done.set()

    pyxll.schedule_call(_do_reload)
    if not done.wait(timeout=30):
        _log.warning("pyxll_reload_module(%s): timed out", module_name)
        return "ERROR: reload timed out (Excel may be busy)", True
    text = result[0]
    return text, text != "OK"


register(
    "pyxll_reload_module",
    (
        "Reload a single Python module from the workspace and rebind all @xl_func "
        "functions with Excel. This reloads only the specified module — it does NOT "
        "perform a full PyXLL add-in reload (which would close the task pane). "
        "Call this after every edit to pyxll_claude_functions.py. "
        "Returns 'OK' on success or a Python traceback on error."
    ),
    {
        "type": "object",
        "properties": {
            "module": {
                "type": "string",
                "description": (
                    "Name of the module to reload. "
                    "Defaults to 'pyxll_claude_functions' if omitted."
                ),
            },
        },
        "required": [],
    },
    _impl,
)
