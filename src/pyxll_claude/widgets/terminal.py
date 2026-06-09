"""
ClaudeTerminalWidget — xterm.js terminal backed by a Claude CLI subprocess.

Hosts a WebViewClient (child-process xterm.js view) and a ClaudeProcess
(Windows ConPTY wrapping the claude CLI).  Starts the MCP server and the
claude CLI once xterm.js reports its initial size, then emits workspace_ready
so the parent pane can initialise dependent widgets.
"""
import base64
import configparser
import json
import logging
import threading
from pathlib import Path

from PySide6.QtCore import Signal, Slot
from PySide6.QtWidgets import QVBoxLayout, QWidget

from pyxll import get_config

from ..claude_process import ClaudeProcess
from ..mcp_server import PyXLLMCPServer, find_free_port
from ..webview import WebViewClient
from ..workspace import ensure_workspace_initialized

_log = logging.getLogger(__name__)

_RESOURCES = Path(__file__).parent.parent / "resources"

# ANSI escape helpers used in status messages written to the terminal.
_R = "\x1b[0m"
_B = "\x1b[1m"
_Y = "\x1b[33m"
_C = "\x1b[36m"


# ---------------------------------------------------------------------------
# Config helpers
# ---------------------------------------------------------------------------

def _get_workspace() -> Path | None:
    """Return the workspace path from pyxll.cfg [CLAUDE] → workspace, or None."""
    try:
        value = get_config().get("CLAUDE", "workspace").strip()
        return Path(value) if value else None
    except (configparser.NoSectionError, configparser.NoOptionError):
        return None
    except Exception:
        _log.warning("Failed to read [CLAUDE] workspace from pyxll.cfg", exc_info=True)
        return None


def _get_mcp_enabled() -> bool:
    """Return False if [CLAUDE] mcp_enabled = false in pyxll.cfg, True otherwise."""
    try:
        value = get_config().get("CLAUDE", "mcp_enabled").strip().lower()
        return value not in ("false", "0", "no", "off")
    except Exception:
        return True


def _get_mcp_port_range() -> tuple[int, int]:
    """Return (start, end) port range from pyxll.cfg [CLAUDE] → mcp_port_range."""
    try:
        value = get_config().get("CLAUDE", "mcp_port_range").strip()
        start, _, end = value.partition("-")
        return int(start.strip()), int(end.strip())
    except Exception:
        return 54717, 54816


# ---------------------------------------------------------------------------
# Widget
# ---------------------------------------------------------------------------

class ClaudeTerminalWidget(QWidget):
    """xterm.js terminal + Claude CLI process, hosted in a child process web view.

    Signals:
        workspace_ready(str)  — workspace path; emitted on the main thread once
                                the Claude CLI has started successfully.
    """
    workspace_ready = Signal(str)    # public: workspace confirmed, Claude started
    _term_write_sig = Signal(bytes)  # private: marshal background → main thread

    def __init__(self, xl_hwnd: int = 0, parent=None):
        super().__init__(parent)
        self._mcp_server   = None
        self._is_destroyed = False

        self._client = WebViewClient(
            url=str(_RESOURCES / "terminal.html"),
            xl_hwnd=xl_hwnd,
            loading_text="Starting Claude terminal…",
            parent=self,
        )
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self._client)

        self._process = ClaudeProcess(self)
        self._client.message_received.connect(self._on_message)
        self._process.data_received.connect(self._send_output)
        self._term_write_sig.connect(self._send_output)
        self.destroyed.connect(self._cleanup)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def write(self, text: str) -> None:
        """Write ANSI/plain text to xterm.js. Thread-safe."""
        self._term_write_sig.emit(text.encode("utf-8"))

    # ------------------------------------------------------------------
    # Internal — message routing
    # ------------------------------------------------------------------

    @Slot(dict)
    def _on_message(self, msg: dict) -> None:
        t = msg.get("type")
        if t == "start_process":
            self._on_start_process(msg["cols"], msg["rows"])
        elif t == "terminal_input":
            self._process.write(msg["text"].encode("utf-8"))
        elif t == "terminal_resize":
            self._process.resize(msg["cols"], msg["rows"])

    def _send_output(self, data: bytes) -> None:
        self._client.send_message({
            "type": "terminal_output",
            "data": base64.b64encode(data).decode(),
        })

    # ------------------------------------------------------------------
    # Internal — process startup
    # ------------------------------------------------------------------

    def _on_start_process(self, cols: int, rows: int) -> None:
        # xterm.js has rendered and posted start_process.  Chromium's render
        # widget exists now; attach all its threads to Excel's input queue so
        # Chromium's internal SetFocus calls succeed and WM_KEYDOWN reaches JS.
        self._client.attach_child_threads()
        self._client.focus_child()

        workspace = _get_workspace()
        if workspace is None:
            self.write(
                f"\r\n{_Y}{_B}Claude workspace not configured.{_R}\r\n\r\n"
                f"Add the following to your {_B}pyxll.cfg{_R} and reload PyXLL:\r\n\r\n"
                f"  {_C}[CLAUDE]\r\n"
                f"  workspace = C:\\path\\to\\your\\workspace{_R}\r\n"
            )
            return
        if not workspace.exists():
            self.write(
                f"\r\n{_Y}{_B}Claude workspace folder not found.{_R}\r\n\r\n"
                f"Configured path:\r\n  {_C}{workspace}{_R}\r\n\r\n"
                f"Create the folder or update {_B}pyxll.cfg{_R}, then reload PyXLL.\r\n"
            )
            return
        warnings = ensure_workspace_initialized(workspace)
        for warning in warnings:
            self.write(f"\r\n{_Y}⚠  {warning}{_R}\r\n")
        threading.Thread(
            target=self._start_mcp_and_claude,
            args=(workspace, cols, rows),
            daemon=True,
            name="pyxll-mcp-start",
        ).start()

    def _start_mcp_and_claude(self, workspace: Path, cols: int, rows: int) -> None:
        """Background thread: start MCP server then the claude CLI."""
        if self._is_destroyed:
            return

        if _get_mcp_enabled():
            port_start, port_end = _get_mcp_port_range()
            port = find_free_port(port_start, port_end)
            if port is not None:
                if self._is_destroyed:
                    return
                server = PyXLLMCPServer(workspace=workspace, port=port)
                if server.start():
                    if self._is_destroyed:
                        server.stop()
                        return
                    self._mcp_server = server
                    mcp_json = workspace / ".mcp.json"
                    mcp_json.write_text(
                        json.dumps(
                            {"mcpServers": {"pyxll": {
                                "type": "sse",
                                "url": f"http://localhost:{port}/sse",
                            }}},
                            indent=2,
                        ),
                        encoding="utf-8",
                    )
                else:
                    _log.warning("MCP server failed to bind port %d", port)
            else:
                _log.warning("No free MCP port in range %d–%d", port_start, port_end)

        self._process.start(cols=cols, rows=rows, workspace=workspace)
        self.workspace_ready.emit(str(workspace))  # queued → main thread

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------

    def _cleanup(self, _obj=None) -> None:
        if self._is_destroyed:
            return
        self._is_destroyed = True

        # Stop the Claude CLI first so it disconnects from the MCP server
        # before the server stops.
        if hasattr(self, "_process") and self._process:
            self._process.terminate()
            if not self._process.wait(5.0):
                _log.warning("ClaudeProcess failed to stop gracefully")

        if self._mcp_server:
            self._mcp_server.stop()
            self._mcp_server = None

    def closeEvent(self, event) -> None:
        self._cleanup()
        self._client.close()
        super().closeEvent(event)
