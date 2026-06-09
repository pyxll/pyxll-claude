"""
PySide6 Custom Task Pane hosting the Claude AI terminal.

Layout — a QTabWidget with two tabs:

  Terminal   — the xterm.js QWebEngineView connected to the claude PTY.
  Functions  — Monaco (VS Code) editor for pyxll_claude_functions.py in the workspace.
               Tab title shows "Functions *" when there are unsaved edits.  Ctrl+S saves
               the file and automatically reloads functions via pyxll.rebind().
               The Load Functions button below is a manual fallback.

Startup sequence:
  1. ClaudeTerminalWidget opens; QWebEngineView loads terminal.html.
  2. xterm.js calls bridge.startProcess(cols, rows).
  3. TerminalBridge validates the [CLAUDE] workspace from pyxll.cfg, shows
     an error in the terminal if misconfigured, or:
     a. Initialises the workspace (CLAUDE.md, skill, pyxll_claude_functions.py).
     b. Starts ClaudeProcess in that directory.
     c. Emits workspaceReady so the Functions tab populates.
"""
import base64
import configparser
import importlib
import json
import logging
import sys
import threading
from pathlib import Path

from PySide6.QtCore import QFileSystemWatcher, QObject, QTimer, QUrl, Signal, Slot
from PySide6.QtWebChannel import QWebChannel
from PySide6.QtWebEngineCore import QWebEngineSettings
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWidgets import (
    QApplication, QLabel, QPushButton, QTabWidget,
    QVBoxLayout, QWidget,
)

from pyxll import create_ctp, CTPDockPositionRight, get_config, rebind

from .mcp_server import PyXLLMCPServer, find_free_port
from .terminal import ClaudeProcess
from .workspace import ensure_workspace_initialized

_log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Config helper
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


# ANSI helpers for terminal messages
_R = "\x1b[0m"
_B = "\x1b[1m"
_Y = "\x1b[33m"
_C = "\x1b[36m"


# ---------------------------------------------------------------------------
# Qt/JS bridges
# ---------------------------------------------------------------------------

class CodeEditorBridge(QObject):
    """QWebChannel bridge registered as 'editorBridge' in the Monaco editor page.

    Signals (Python → JS):
        loadContent(str)  — set editor text and update the saved baseline
        markSaved()       — advance the saved baseline to the current editor text

    Slots (JS → Python):
        editorReady()          — Monaco has initialised; flush any buffered content
        setDirty(bool)         — editor dirty-state changed
        saveContent(str)       — user pressed Ctrl+S; text is the current editor value
    """

    loadContent  = Signal(str)
    markSaved    = Signal()
    requestSave  = Signal()   # → JS: ask the editor to call saveContent with current text

    def __init__(self, parent=None):
        super().__init__(parent)
        self._ready           = False
        self._pending_content = None
        self.on_dirty_changed = None   # callable(bool)
        self.on_save          = None   # callable(str)

    def set_content(self, text: str):
        """Push new content to the editor from Python."""
        if self._ready:
            self.loadContent.emit(text)
        else:
            self._pending_content = text

    @Slot()
    def editorReady(self):
        self._ready = True
        if self._pending_content is not None:
            self.loadContent.emit(self._pending_content)
            self._pending_content = None

    @Slot(bool)
    def setDirty(self, dirty: bool):
        if self.on_dirty_changed:
            self.on_dirty_changed(dirty)

    @Slot(str)
    def saveContent(self, text: str):
        if self.on_save:
            self.on_save(text)


class TerminalBridge(QObject):
    """QWebChannel bridge registered as 'bridge' in the xterm.js page.

    Slots (called from JavaScript):
        startProcess(cols, rows)   — validate workspace, initialise it, start claude
        receiveInput(text)         — forward keystrokes to the PTY
        resizeTerminal(cols, rows) — forward terminal resize to the PTY
        copyToClipboard(text)      — write text to the system clipboard
        pasteFromClipboard()       — return current system clipboard text

    Signals (emitted to JavaScript or connected within Python):
        outputReady(str)    — base64-encoded PTY output chunk → xterm.js
        workspaceReady(str) — workspace path string, emitted after process starts
    """

    outputReady    = Signal(str)
    workspaceReady = Signal(str)

    def __init__(self, process: ClaudeProcess, parent=None):
        super().__init__(parent)
        self._process = process
        self._mcp_server = None
        process.data_received.connect(self._relay_output)

    @Slot(int, int)
    def startProcess(self, cols: int, rows: int):
        """Validate the workspace and initialise it, then hand off to a background
        thread for socket and process operations to avoid blocking Excel's main thread.
        """
        workspace = _get_workspace()

        if workspace is None:
            self._write(
                f"\r\n{_Y}{_B}Claude workspace not configured.{_R}\r\n\r\n"
                f"Add the following to your {_B}pyxll.cfg{_R} and reload PyXLL:\r\n\r\n"
                f"  {_C}[CLAUDE]\r\n"
                f"  workspace = C:\\path\\to\\your\\workspace{_R}\r\n"
            )
            return

        if not workspace.exists():
            self._write(
                f"\r\n{_Y}{_B}Claude workspace folder not found.{_R}\r\n\r\n"
                f"Configured path:\r\n  {_C}{workspace}{_R}\r\n\r\n"
                f"Create the folder or update {_B}pyxll.cfg{_R}, then reload PyXLL.\r\n"
            )
            return

        warnings = ensure_workspace_initialized(workspace)
        for warning in warnings:
            self._write(f"\r\n{_Y}⚠  {warning}{_R}\r\n")

        threading.Thread(
            target=self._start_mcp_and_claude,
            args=(workspace, cols, rows),
            daemon=True,
            name="pyxll-mcp-start",
        ).start()

    def _start_mcp_and_claude(self, workspace: Path, cols: int, rows: int) -> None:
        """Background thread: find a free port, start the MCP server, then claude."""
        if _get_mcp_enabled():
            port_start, port_end = _get_mcp_port_range()
            port = find_free_port(port_start, port_end)
            if port is not None:
                server = PyXLLMCPServer(workspace=workspace, port=port)
                if server.start():
                    self._mcp_server = server
                    # Write .mcp.json to the workspace root — Claude Code reads this
                    # file from the CWD at startup. Must be in the project root, not
                    # .claude/; adding mcpServers to settings.json causes errors.
                    mcp_json_path = workspace / ".mcp.json"
                    mcp_json_path.write_text(
                        json.dumps(
                            {"mcpServers": {"pyxll": {"type": "sse", "url": f"http://localhost:{port}/sse"}}},
                            indent=2,
                        ),
                        encoding="utf-8",
                    )
                else:
                    self._mcp_server = None
                    _log.warning("MCP server failed to bind port %d", port)
            else:
                _log.warning("No free MCP port found in range %d-%d", port_start, port_end)

        self._process.start(cols=cols, rows=rows, workspace=workspace)
        self.workspaceReady.emit(str(workspace))

    def stop_mcp_server(self) -> None:
        """Stop the MCP server if running."""
        if self._mcp_server is not None:
            self._mcp_server.stop()
            self._mcp_server = None

    @Slot(str)
    def receiveInput(self, text: str):
        """Forward keyboard/paste input from xterm.js to the claude PTY."""
        self._process.write(text.encode("utf-8"))

    @Slot(int, int)
    def resizeTerminal(self, cols: int, rows: int):
        """Notify the PTY of a terminal size change."""
        self._process.resize(cols, rows)

    @Slot(str)
    def copyToClipboard(self, text: str):
        """Write text to the system clipboard."""
        QApplication.clipboard().setText(text)

    @Slot(result=str)
    def pasteFromClipboard(self) -> str:
        """Return the current system clipboard text."""
        return QApplication.clipboard().text()

    def _relay_output(self, data: bytes):
        self.outputReady.emit(base64.b64encode(data).decode("ascii"))

    def _write(self, text: str):
        """Write an ANSI/plain-text message directly to xterm.js."""
        self.outputReady.emit(base64.b64encode(text.encode("utf-8")).decode("ascii"))


# ---------------------------------------------------------------------------
# Main widget
# ---------------------------------------------------------------------------

class ClaudeTerminalWidget(QWidget):
    """Custom Task Pane: Terminal tab + Functions tab."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._process   = ClaudeProcess(self)
        self._workspace = None
        self._watcher   = None
        self._skip_next_watcher = False
        self._editor_dirty      = False
        # These are set by _build_functions_tab() and used by later methods.
        self._editor_bridge = None
        self._status_label  = None
        self._build_ui()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self._tabs = QTabWidget(self)
        self._tabs.addTab(self._build_terminal_tab(),   "Terminal")
        self._tabs.addTab(self._build_functions_tab(),  "Functions")
        layout.addWidget(self._tabs, stretch=1)

        # Load Functions button and status are always visible below the tabs.
        bottom = QWidget(self)
        bottom_layout = QVBoxLayout(bottom)
        bottom_layout.setContentsMargins(8, 6, 8, 6)
        bottom_layout.setSpacing(4)

        load_btn = QPushButton("Load Functions", bottom)
        load_btn.clicked.connect(self._on_load_functions)
        bottom_layout.addWidget(load_btn)

        self._status_label = QLabel(bottom)
        self._status_label.setWordWrap(True)
        bottom_layout.addWidget(self._status_label)

        layout.addWidget(bottom)

    def _build_terminal_tab(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)

        self._view = QWebEngineView(container)

        settings = self._view.page().settings()
        settings.setAttribute(
            QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls, True
        )

        self._channel = QWebChannel(self)
        self._bridge  = TerminalBridge(self._process, self)
        self._bridge.workspaceReady.connect(self._on_workspace_ready)
        self._channel.registerObject("bridge", self._bridge)
        self._view.page().setWebChannel(self._channel)

        html_path = Path(__file__).parent / "resources" / "terminal.html"
        self._view.setUrl(QUrl.fromLocalFile(str(html_path)))

        layout.addWidget(self._view)
        return container

    def _build_functions_tab(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)

        view = QWebEngineView(container)
        settings = view.page().settings()
        settings.setAttribute(
            QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls, True
        )

        channel = QWebChannel(self)
        self._editor_bridge = CodeEditorBridge(self)
        self._editor_bridge.on_dirty_changed = self._on_editor_dirty_changed
        self._editor_bridge.on_save          = self._on_editor_save
        channel.registerObject("editorBridge", self._editor_bridge)
        view.page().setWebChannel(channel)

        html_path = Path(__file__).parent / "resources" / "editor.html"
        view.setUrl(QUrl.fromLocalFile(str(html_path)))

        layout.addWidget(view)
        return container

    # ------------------------------------------------------------------
    # Workspace ready — populate Functions tab and start file watcher
    # ------------------------------------------------------------------

    @Slot(str)
    def _on_workspace_ready(self, workspace_str: str):
        self._workspace = Path(workspace_str)
        self._refresh_functions_view()

        xl_path = str(self._workspace / "pyxll_claude_functions.py")
        self._watcher = QFileSystemWatcher([xl_path], self)
        self._watcher.fileChanged.connect(self._on_pyxll_claude_functions_changed)

    def _on_pyxll_claude_functions_changed(self, path: str):
        # Re-add the path so the watcher survives editors that replace the file.
        if self._watcher and path not in self._watcher.files():
            self._watcher.addPath(path)
        # Skip the event that fires immediately after our own save.
        if self._skip_next_watcher:
            self._skip_next_watcher = False
            return
        self._refresh_functions_view()

    def _refresh_functions_view(self):
        if self._workspace is None or self._editor_bridge is None:
            return
        xl_path = self._workspace / "pyxll_claude_functions.py"
        try:
            self._editor_bridge.set_content(xl_path.read_text(encoding="utf-8"))
        except Exception as exc:
            self._editor_bridge.set_content(f"# Error reading {xl_path}:\n# {exc}")

    # ------------------------------------------------------------------
    # Editor dirty-state and save
    # ------------------------------------------------------------------

    def _on_editor_dirty_changed(self, dirty: bool):
        self._editor_dirty = dirty
        self._tabs.setTabText(1, "Functions *" if dirty else "Functions")

    def _on_editor_save(self, text: str):
        if self._workspace is None:
            self._set_status("Workspace not ready.", ok=False)
            return
        xl_path = self._workspace / "pyxll_claude_functions.py"
        try:
            self._skip_next_watcher = True
            xl_path.write_text(text, encoding="utf-8")
            self._editor_bridge.markSaved.emit()
        except Exception as exc:
            self._skip_next_watcher = False
            self._set_status(f"Save failed: {exc}", ok=False)
            return
        self._on_load_functions()

    # ------------------------------------------------------------------
    # Load Functions button
    # ------------------------------------------------------------------

    def _on_load_functions(self):
        if self._workspace is None:
            self._set_status("Workspace not ready.", ok=False)
            return

        # If the editor has unsaved changes, trigger a save first.  The save
        # path calls _on_load_functions() again once the file is written.
        if self._editor_dirty and self._editor_bridge is not None:
            self._editor_bridge.requestSave.emit()
            return

        ws_str = str(self._workspace)
        if ws_str not in sys.path:
            sys.path.insert(0, ws_str)

        try:
            if "pyxll_claude_functions" in sys.modules:
                importlib.reload(sys.modules["pyxll_claude_functions"])
            else:
                importlib.import_module("pyxll_claude_functions")

            rebind()
            self._set_status("Functions loaded successfully.", ok=True)
            QTimer.singleShot(5000, self._clear_status)
        except Exception as exc:
            _log.error("Error loading pyxll_claude_functions", exc_info=True)
            self._set_status(str(exc), ok=False)

    def _set_status(self, message: str, ok: bool):
        if self._status_label is None:
            return
        colour = "#4ec9b0" if ok else "#f44747"
        self._status_label.setStyleSheet(f"color: {colour};")
        self._status_label.setText(message)

    def _clear_status(self):
        if self._status_label is not None:
            self._status_label.setText("")

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------

    def closeEvent(self, event):
        self._bridge.stop_mcp_server()
        self._process.terminate()
        super().closeEvent(event)


# ---------------------------------------------------------------------------
# Entry point called by the ribbon button
# ---------------------------------------------------------------------------

def show_claude_pane():
    """Create and display the Claude AI custom task pane in Excel."""
    app = QApplication.instance()
    if app is None:
        app = QApplication([])

    ctp_widget = ClaudeTerminalWidget()
    create_ctp(ctp_widget, width=800, position=CTPDockPositionRight)
