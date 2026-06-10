"""
CodeEditorWidget — Monaco editor backed by a child-process web view.

Hosts a WebViewClient (child-process Monaco view), watches the workspace
functions file for external changes, and relays save/dirty events between
the Monaco JS and the parent pane.
"""
import logging
from pathlib import Path

from PySide6.QtCore import QFileSystemWatcher, Signal, Slot
from PySide6.QtWidgets import QVBoxLayout, QWidget

from ..webview import WebViewClient

_log = logging.getLogger(__name__)

_RESOURCES = Path(__file__).parent.parent / "resources"


class CodeEditorWidget(QWidget):
    """Monaco editor for pyxll_claude_functions.py, hosted in a child process.

    Signals:
        dirty_changed(bool)  — editor dirty state changed
        saved()              — file written successfully; proceed to load functions
        save_failed(str)     — save error message
    """
    dirty_changed = Signal(bool)
    saved         = Signal()
    save_failed   = Signal(str)

    def __init__(self, xl_hwnd: int = 0, parent=None):
        super().__init__(parent)
        self._workspace         = None
        self._watcher           = None
        self._skip_next_watcher = False
        self._is_dirty          = False
        self._ready             = False
        self._pending           = None   # content queued before editor_ready

        self._client = WebViewClient(
            url=str(_RESOURCES / "editor.html"),
            xl_hwnd=xl_hwnd,
            loading_text="Loading editor…",
            parent=self,
        )
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self._client)

        self._client.message_received.connect(self._on_message)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    @property
    def is_dirty(self) -> bool:
        return self._is_dirty

    @property
    def workspace(self) -> Path | None:
        return self._workspace

    def load_workspace(self, workspace_str: str) -> None:
        """Call when the workspace is confirmed ready (connected to terminal's workspace_ready)."""
        self._workspace = Path(workspace_str)
        self._push_content()
        xl_path = str(self._workspace / "pyxll_claude_functions.py")
        self._watcher = QFileSystemWatcher([xl_path], self)
        self._watcher.fileChanged.connect(self._on_file_changed)

    def request_save(self) -> None:
        """Ask the editor to save its current content (used when Load Functions is clicked while dirty)."""
        self._client.send_message({"type": "editor_request_save"})

    # ------------------------------------------------------------------
    # Internal — message routing
    # ------------------------------------------------------------------

    @Slot(dict)
    def _on_message(self, msg: dict) -> None:
        t = msg.get("type")
        if t == "editor_ready":
            self._on_editor_ready()
        elif t == "editor_dirty":
            self._on_dirty(msg["dirty"])
        elif t == "editor_save":
            self._on_save(msg["text"])
        elif t == "editor_selection_changed":
            self._client.refocus_renderer()

    def _push_content(self) -> None:
        if self._workspace is None:
            return
        xl_path = self._workspace / "pyxll_claude_functions.py"
        try:
            text = xl_path.read_text(encoding="utf-8")
        except Exception as exc:
            text = f"# Error reading {xl_path}:\n# {exc}"
        if self._ready:
            self._client.send_message({"type": "editor_load", "text": text})
        else:
            self._pending = text

    def _on_editor_ready(self) -> None:
        self._ready = True
        self._client.attach_child_threads()
        if self._pending is not None:
            self._client.send_message({"type": "editor_load", "text": self._pending})
            self._pending = None
        else:
            self._push_content()

    def _on_dirty(self, dirty: bool) -> None:
        self._is_dirty = dirty
        self.dirty_changed.emit(dirty)

    def _on_save(self, text: str) -> None:
        if self._workspace is None:
            self.save_failed.emit("Workspace not ready.")
            return
        xl_path = self._workspace / "pyxll_claude_functions.py"
        try:
            self._skip_next_watcher = True
            xl_path.write_text(text, encoding="utf-8")
            self._client.send_message({"type": "editor_mark_saved"})
        except Exception as exc:
            self._skip_next_watcher = False
            self.save_failed.emit(f"Save failed: {exc}")
            return
        # Update dirty state immediately so the caller sees is_dirty=False
        # when responding to saved(); JS will confirm via editor_dirty:false shortly.
        self._is_dirty = False
        self.saved.emit()

    def _on_file_changed(self, path: str) -> None:
        if self._watcher and path not in self._watcher.files():
            self._watcher.addPath(path)
        if self._skip_next_watcher:
            self._skip_next_watcher = False
            return
        self._push_content()

    def closeEvent(self, event) -> None:
        self._client.close()
        super().closeEvent(event)
