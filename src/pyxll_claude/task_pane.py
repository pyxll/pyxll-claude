"""
PySide6 Custom Task Pane hosting the Claude AI terminal and code editor.

ClaudeCustomTaskPane is a QTabWidget container that hosts two child widgets:
  - ClaudeTerminalWidget  (Terminal tab)  — xterm.js + Claude CLI subprocess
  - CodeEditorWidget      (Functions tab) — Monaco editor for pyxll_claude_functions.py

Each widget runs its web view in a separate child process to isolate Chromium's
message pump from Excel's COM message pump.

Layout:
  Terminal | Functions      ← QTabWidget tab bar
  ┌──────────────────────┐
  │  terminal / editor   │  ← active child-process web view
  └──────────────────────┘
  [Load Functions]  <status> ← bottom bar
"""
import importlib
import logging
import sys
import weakref

from PySide6.QtCore import QTimer, Slot
from PySide6.QtWidgets import (
    QApplication, QLabel, QPushButton, QTabWidget,
    QVBoxLayout, QWidget,
)

from pyxll import create_ctp, CTPDockPositionRight, rebind, xl_app

from .widgets import CodeEditorWidget, ClaudeTerminalWidget

_log = logging.getLogger(__name__)


class ClaudeCustomTaskPane(QWidget):
    """Custom Task Pane: Terminal + Functions tabs."""

    def __init__(self, parent=None):
        super().__init__(parent)
        xl_hwnd = 0
        try:
            xl_hwnd = xl_app().Hwnd
        except Exception:
            _log.warning("Could not read Excel Application.Hwnd", exc_info=True)

        self._build_ui(xl_hwnd)
        self._connect_signals()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self, xl_hwnd: int) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        self._tabs = QTabWidget(self)

        self._terminal = ClaudeTerminalWidget(xl_hwnd=xl_hwnd, parent=self)
        self._editor   = CodeEditorWidget(xl_hwnd=xl_hwnd, parent=self)

        self._tabs.addTab(self._terminal, "Terminal")
        self._tabs.addTab(self._editor,   "Functions")
        root.addWidget(self._tabs, stretch=1)

        bottom = QWidget(self)
        bl = QVBoxLayout(bottom)
        bl.setContentsMargins(8, 6, 8, 6)
        bl.setSpacing(4)
        load_btn = QPushButton("Load Functions", bottom)
        load_btn.clicked.connect(self._on_load_functions)
        bl.addWidget(load_btn)
        self._status_label = QLabel(bottom)
        self._status_label.setWordWrap(True)
        bl.addWidget(self._status_label)
        root.addWidget(bottom)

    # ------------------------------------------------------------------
    # Signal wiring
    # ------------------------------------------------------------------

    def _connect_signals(self) -> None:
        self._terminal.workspace_ready.connect(self._editor.load_workspace)
        self._editor.dirty_changed.connect(self._on_editor_dirty)
        self._editor.saved.connect(self._on_load_functions)
        self._editor.save_failed.connect(lambda msg: self._set_status(msg, ok=False))

    def _on_editor_dirty(self, dirty: bool) -> None:
        self._tabs.setTabText(1, "Functions *" if dirty else "Functions")

    # ------------------------------------------------------------------
    # Load Functions
    # ------------------------------------------------------------------

    @Slot()
    def _on_load_functions(self) -> None:
        workspace = self._editor.workspace
        if workspace is None:
            self._set_status("Workspace not ready.", ok=False)
            return
        if self._editor.is_dirty:
            self._editor.request_save()
            return
        ws_str = str(workspace)
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

    def _set_status(self, message: str, ok: bool) -> None:
        self._status_label.setStyleSheet(
            f"color: {'#4ec9b0' if ok else '#f44747'};"
        )
        self._status_label.setText(message)

    def _clear_status(self) -> None:
        self._status_label.setText("")

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------

    def closeEvent(self, event) -> None:
        # Qt won't call closeEvent on child widgets, so explicitly close them
        # in order: terminal first (Claude → MCP → Chrome), then editor.
        self._terminal.close()
        self._editor.close()
        super().closeEvent(event)


# ---------------------------------------------------------------------------
# Entry point called by the ribbon button
# ---------------------------------------------------------------------------

_active_pane: "weakref.ref[ClaudeCustomTaskPane] | None" = None


def show_claude_pane():
    """Create and display the Claude AI custom task pane in Excel."""
    global _active_pane
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    ctp_widget = ClaudeCustomTaskPane()
    create_ctp(ctp_widget, width=800, position=CTPDockPositionRight)
    _active_pane = weakref.ref(ctp_widget)


def is_claude_pane_open() -> bool:
    """Return True if the Claude terminal task pane is currently open."""
    if _active_pane is None:
        return False
    pane = _active_pane()
    return pane is not None and pane.isVisible()
