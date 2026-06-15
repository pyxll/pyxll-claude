import traceback

import pyxll

from . import register
from ._utils import run_on_main_thread


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    try:
        run_on_main_thread(pyxll.reload)
        return (
            "PyXLL add-in reload has been scheduled. "
            "Excel will reload the add-in shortly. "
            "WARNING: all custom task panes — including this Claude terminal — "
            "will be closed by the reload. The MCP connection will be lost. "
            "You will need to reopen the Claude terminal pane manually after the reload completes.",
            False,
        )
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


register(
    "pyxll_reload_addin",
    (
        "Reload the entire PyXLL add-in (equivalent to clicking 'Reload PyXLL' in the ribbon). "
        "Use this when changes require a full add-in reload — for example, after editing "
        "pyxll.cfg, adding a new module to the modules list, or making changes to package "
        "structure that pyxll_reload_module cannot pick up. "
        "Before calling this tool, call pyxll_is_terminal_open. "
        "If it returns True, you MUST stop, tell the user the terminal will close and "
        "ask them to confirm before proceeding — do NOT call this tool until the user "
        "explicitly agrees. "
        "If it returns False the terminal is already closed and you may proceed. "
        "The MCP server itself survives the reload — only the terminal UI is affected. "
        "Prefer pyxll_reload_module for everyday function edits."
    ),
    {"type": "object", "properties": {}, "required": []},
    _impl,
)
