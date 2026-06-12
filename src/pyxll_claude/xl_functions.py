"""
PyXLL-decorated functions and ribbon action callbacks for pyxll-claude.

This module is returned by pyxll_modules() and loaded by PyXLL automatically.
Ribbon onAction/getImage/getContent callbacks must match ribbon.xml attributes.
"""
import logging

import pyxll
import win32api
from pyxll import xl_on_open, xl_menu, xl_app
from PySide6.QtWidgets import QApplication, QMessageBox

from .mcp_server import (
    get_global_port, get_or_start_global, set_on_server_change,
    stop_global, write_mcp_json,
)
from .task_pane import show_claude_pane
from .workspace import ensure_workspace_initialized, get_workspace_path

_log = logging.getLogger(__name__)

_MCP_NS = "http://schemas.microsoft.com/office/2009/07/customui"


def _get_mcp_port_range() -> tuple[int, int]:
    try:
        from pyxll import get_config
        value = get_config().get("CLAUDE", "mcp_port_range").strip()
        start, _, end = value.partition("-")
        return int(start.strip()), int(end.strip())
    except Exception:
        return 54717, 54816


def _is_mcp_running() -> bool:
    return get_global_port() is not None


# ---------------------------------------------------------------------------
# Ribbon lifecycle
# ---------------------------------------------------------------------------

def on_ribbon_load(ribbon_ui) -> None:
    """Capture IRibbonUI on load; wire it up as the server-change callback."""
    def _invalidate():
        try:
            ribbon_ui.InvalidateControl("PyXLLMCPServer")
        except Exception:
            _log.debug("ribbon invalidate failed", exc_info=True)

    set_on_server_change(_invalidate)


# ---------------------------------------------------------------------------
# Dynamic MCP button callbacks
# ---------------------------------------------------------------------------

def get_mcp_button_image(control):
    name = "pyxll_claude.resources:mcp_stop.png" if _is_mcp_running() else "pyxll_claude.resources:mcp_start.png"
    return pyxll.load_image(name)


def get_mcp_button_screentip(control) -> str:
    if _is_mcp_running():
        return f"PyXLL MCP server running on port {get_global_port()}"
    return "Start PyXLL MCP Server"


def get_mcp_menu_content(control) -> str:
    if _is_mcp_running():
        return (
            f'<menu xmlns="{_MCP_NS}">'
            f'  <button id="McpShowEndpoint"'
            f'          label="Show Endpoint"'
            f'          onAction="pyxll_claude.xl_functions.show_mcp_endpoint"/>'
            f'  <menuSeparator id="McpSep"/>'
            f'  <button id="McpStop"'
            f'          label="Stop MCP Server"'
            f'          image="pyxll_claude.resources:mcp_stop.png"'
            f'          onAction="pyxll_claude.xl_functions.stop_mcp_server"/>'
            f'</menu>'
        )
    return (
        f'<menu xmlns="{_MCP_NS}">'
        f'  <button id="McpStart"'
        f'          label="Start MCP Server"'
        f'          image="pyxll_claude.resources:mcp_start.png"'
        f'          onAction="pyxll_claude.xl_functions.start_mcp_server"/>'
        f'</menu>'
    )


# ---------------------------------------------------------------------------
# MCP server actions
# ---------------------------------------------------------------------------

def _show_endpoint_dialog(port: int) -> None:
    """QMessageBox with a Copy URL button for the MCP endpoint."""
    url = f"http://localhost:{port}/mcp"
    app = QApplication.instance() or QApplication([])
    box = QMessageBox()
    box.setWindowTitle("PyXLL MCP Server")
    box.setText(
        f"PyXLL MCP server is running.\n\n"
        f"Endpoint:  {url}\n\n"
        f"Add this endpoint to your Claude session's MCP settings "
        f"to give Claude access to this Excel workbook."
    )
    box.setIcon(QMessageBox.Icon.Information)
    copy_btn = box.addButton("Copy URL", QMessageBox.ButtonRole.ActionRole)
    ok_btn = box.addButton(QMessageBox.StandardButton.Ok)
    box.setDefaultButton(ok_btn)
    box.exec()
    if box.clickedButton() is copy_btn:
        app.clipboard().setText(url)


def start_mcp_server(control=None) -> None:
    """Start the MCP server and show its endpoint."""
    try:
        workspace = get_workspace_path()
        ensure_workspace_initialized(workspace)
        port_start, port_end = _get_mcp_port_range()
        result = get_or_start_global(workspace, port_start, port_end)
        if result is None:
            try:
                hwnd = xl_app().Hwnd
            except Exception:
                hwnd = 0
            win32api.MessageBox(
                hwnd,
                "Could not start the PyXLL MCP server.\n\n"
                "No free port was found in the configured range.",
                "PyXLL MCP Server",
                0x10,  # MB_ICONERROR
            )
            return
        _, port = result
        write_mcp_json(workspace, port)
        _show_endpoint_dialog(port)
    except Exception:
        _log.error("Failed to start MCP server", exc_info=True)


def stop_mcp_server(control=None) -> None:
    """Stop the running MCP server."""
    try:
        stop_global()
    except Exception:
        _log.error("Failed to stop MCP server", exc_info=True)


def show_mcp_endpoint(control=None) -> None:
    """Show the current MCP server endpoint."""
    try:
        port = get_global_port()
        if port is None:
            try:
                hwnd = xl_app().Hwnd
            except Exception:
                hwnd = 0
            win32api.MessageBox(hwnd, "The PyXLL MCP server is not running.", "PyXLL MCP Server", 0x40)
            return
        _show_endpoint_dialog(port)
    except Exception:
        _log.error("Failed to show MCP endpoint", exc_info=True)


# ---------------------------------------------------------------------------
# PyXLL hooks and menu entries
# ---------------------------------------------------------------------------

@xl_on_open
def on_open(import_info):
    """Called by PyXLL after all modules have been imported."""
    _log.info("pyxll_claude loaded successfully.")


def open_claude_pane(control=None):
    """Ribbon button action: open the Claude AI task pane."""
    try:
        show_claude_pane()
    except Exception:
        _log.error("Failed to open Claude task pane", exc_info=True)


@xl_menu("Open Claude Terminal", menu="Claude AI")
def menu_open_claude_pane():
    open_claude_pane()


@xl_menu("MCP Server", menu="Claude AI")
def menu_start_mcp_server():
    start_mcp_server()
