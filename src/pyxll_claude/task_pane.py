"""
PySide6 Custom Task Pane hosting the Claude AI terminal.

Layout — a QTabWidget with two tabs:

  Terminal   — the xterm.js QWebEngineView connected to the claude PTY.
  Functions  — read-only viewer for pyxll_claude_functions.py in the workspace, plus
               a Load Functions button that imports/reloads the module and
               calls pyxll.rebind() to register changes with Excel.

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
import logging
import sys
from pathlib import Path

from PySide6.QtCore import QFileSystemWatcher, QObject, QUrl, Signal, Slot
from PySide6.QtGui import QFont
from PySide6.QtWebChannel import QWebChannel
from PySide6.QtWebEngineCore import QWebEngineSettings
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWidgets import (
    QApplication, QLabel, QPushButton, QTabWidget,
    QTextEdit, QVBoxLayout, QWidget,
)

from pyxll_claude.terminal import ClaudeProcess

_log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Config helper
# ---------------------------------------------------------------------------

def _get_workspace() -> Path | None:
    """Return the workspace path from pyxll.cfg [CLAUDE] → workspace, or None."""
    try:
        from pyxll import get_config
        value = get_config().get("CLAUDE", "workspace").strip()
        return Path(value) if value else None
    except (configparser.NoSectionError, configparser.NoOptionError):
        return None
    except Exception:
        _log.warning("Failed to read [CLAUDE] workspace from pyxll.cfg", exc_info=True)
        return None


# ANSI helpers for terminal messages
_R = "\x1b[0m"
_B = "\x1b[1m"
_Y = "\x1b[33m"
_C = "\x1b[36m"


# ---------------------------------------------------------------------------
# Qt/JS bridge
# ---------------------------------------------------------------------------

class TerminalBridge(QObject):
    """QWebChannel bridge registered as 'bridge' in the xterm.js page.

    Slots (called from JavaScript):
        startProcess(cols, rows)   — validate workspace, initialise it, start claude
        receiveInput(text)         — forward keystrokes to the PTY
        resizeTerminal(cols, rows) — forward terminal resize to the PTY

    Signals (emitted to JavaScript or connected within Python):
        outputReady(str)    — base64-encoded PTY output chunk → xterm.js
        workspaceReady(str) — workspace path string, emitted after process starts
    """

    outputReady    = Signal(str)
    workspaceReady = Signal(str)

    def __init__(self, process: ClaudeProcess, parent=None):
        super().__init__(parent)
        self._process = process
        process.data_received.connect(self._relay_output)

    @Slot(int, int)
    def startProcess(self, cols: int, rows: int):
        """Validate the workspace, initialise it, then start claude."""
        from pyxll_claude.workspace import ensure_workspace_initialized

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

        self._process.start(cols=cols, rows=rows, workspace=workspace)
        self.workspaceReady.emit(str(workspace))

    @Slot(str)
    def receiveInput(self, text: str):
        """Forward keyboard/paste input from xterm.js to the claude PTY."""
        self._process.write(text.encode("utf-8"))

    @Slot(int, int)
    def resizeTerminal(self, cols: int, rows: int):
        """Notify the PTY of a terminal size change."""
        self._process.resize(cols, rows)

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
        # These are set by _build_functions_tab() and used by later methods.
        self._functions_view = None
        self._status_label   = None
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
        layout.setContentsMargins(8, 8, 8, 8)

        self._functions_view = QTextEdit(container)
        self._functions_view.setReadOnly(True)
        font = QFont("Cascadia Code, Consolas, Courier New, monospace")
        font.setPointSize(10)
        self._functions_view.setFont(font)
        self._functions_view.setPlaceholderText(
            "pyxll_claude_functions.py will appear here once the workspace is configured."
        )
        layout.addWidget(self._functions_view)

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
        self._refresh_functions_view()
        # Some editors replace the file rather than modifying it in-place;
        # re-add the path so the watcher survives that.
        if self._watcher and path not in self._watcher.files():
            self._watcher.addPath(path)

    def _refresh_functions_view(self):
        if self._workspace is None or self._functions_view is None:
            return
        xl_path = self._workspace / "pyxll_claude_functions.py"
        try:
            self._functions_view.setPlainText(
                xl_path.read_text(encoding="utf-8")
            )
        except Exception as exc:
            self._functions_view.setPlainText(f"# Error reading {xl_path}:\n# {exc}")

    # ------------------------------------------------------------------
    # Load Functions button
    # ------------------------------------------------------------------

    def _on_load_functions(self):
        if self._workspace is None:
            self._set_status("Workspace not ready.", ok=False)
            return

        ws_str = str(self._workspace)
        if ws_str not in sys.path:
            sys.path.insert(0, ws_str)

        try:
            if "pyxll_claude_functions" in sys.modules:
                importlib.reload(sys.modules["pyxll_claude_functions"])
            else:
                importlib.import_module("pyxll_claude_functions")

            from pyxll import rebind
            rebind()
            self._set_status("Functions loaded successfully.", ok=True)
        except Exception as exc:
            _log.error("Error loading pyxll_claude_functions", exc_info=True)
            self._set_status(str(exc), ok=False)

    def _set_status(self, message: str, ok: bool):
        if self._status_label is None:
            return
        colour = "#4ec9b0" if ok else "#f44747"
        self._status_label.setStyleSheet(f"color: {colour};")
        self._status_label.setText(message)

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------

    def closeEvent(self, event):
        self._process.terminate()
        super().closeEvent(event)


# ---------------------------------------------------------------------------
# Entry point called by the ribbon button
# ---------------------------------------------------------------------------

def show_claude_pane():
    """Create and display the Claude AI custom task pane in Excel."""
    from pyxll import create_ctp, CTPDockPositionRight

    app = QApplication.instance()
    if app is None:
        app = QApplication([])

    ctp_widget = ClaudeTerminalWidget()
    create_ctp(ctp_widget, width=800, position=CTPDockPositionRight)
