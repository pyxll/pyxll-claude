import traceback

from . import register
from ._utils import run_on_main_thread


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    try:
        from pyxll_claude.task_pane import is_claude_pane_open
        is_open = run_on_main_thread(is_claude_pane_open)
        return str(is_open), False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


register(
    "pyxll_is_terminal_open",
    (
        "Return 'True' if the Claude terminal custom task pane is currently open in Excel, "
        "'False' otherwise. Call this before pyxll_reload_addin to decide whether to warn "
        "the user that the terminal will close — no warning is needed if the terminal is "
        "already closed."
    ),
    {"type": "object", "properties": {}, "required": []},
    _impl,
)
