"""
PyXLL-decorated functions and ribbon action callbacks for pyxll-claude.

This module is returned by pyxll_modules() and loaded by PyXLL automatically.
Ribbon onAction callbacks must match the attribute values in ribbon.xml.
"""
import logging

from pyxll import xl_on_open, xl_menu

_log = logging.getLogger(__name__)


@xl_on_open
def on_open(import_info):
    """Called by PyXLL after all modules have been imported."""
    for module_name, module, exc_info in import_info:
        if module is None:
            exc_type, exc_value, _ = exc_info
            _log.error("pyxll_claude: error loading '%s': %s", module_name, exc_value)
    _log.info("pyxll_claude loaded successfully.")


def open_claude_pane(control=None):
    """Ribbon button action: open the Claude AI task pane.

    `control` is the IRibbonControl COM object passed by Excel (may be None).
    """
    try:
        from pyxll_claude.task_pane import show_claude_pane
        show_claude_pane()
    except Exception:
        _log.error("Failed to open Claude task pane", exc_info=True)


@xl_menu("Open Claude Terminal", menu="Claude AI")
def menu_open_claude_pane():
    """Add-ins menu fallback for users without the ribbon."""
    open_claude_pane()
