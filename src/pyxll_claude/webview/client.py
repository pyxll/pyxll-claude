"""
Parent-process client widget for a child-process web view.

WebViewClient manages a child QProcess running child_process.py, embeds the
child window into this widget, and relays messages between the parent
application and the child's JS bridge over a QLocalSocket IPC channel.

Usage:
    client = WebViewClient(url="file:///path/to/page.html", xl_hwnd=excel_hwnd)
    client.message_received.connect(my_handler)
    client.send_message({"type": "some_message", ...})
"""
import ctypes
import json
import logging
import os
import sys
from pathlib import Path

from PySide6.QtCore import QEvent, QObject, QProcess, Qt, Signal, Slot
from PySide6.QtGui import QWindow
from PySide6.QtNetwork import QLocalServer, QLocalSocket
from PySide6.QtWidgets import QApplication, QLabel, QVBoxLayout, QWidget

_log = logging.getLogger(__name__)

_CHILD_SCRIPT = Path(__file__).parent / "child_process.py"

_ipc_counter = 0


def _focus_foreign_hwnd(hwnd: int) -> None:
    """Transfer Windows keyboard focus to an HWND that lives in another process."""
    user32   = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    my_tid    = kernel32.GetCurrentThreadId()
    child_tid = user32.GetWindowThreadProcessId(hwnd, None)
    if child_tid and my_tid != child_tid:
        user32.AttachThreadInput(my_tid, child_tid, True)
    user32.SetFocus(hwnd)


# ---------------------------------------------------------------------------
# IPC — parent side (socket server)
# ---------------------------------------------------------------------------

class _IPC(QObject):
    """QLocalServer that accepts one connection from the child process.

    The 'ready' message (HWND handshake) is handled internally and emitted
    via child_hwnd_ready.  All other messages are emitted via message_received.
    """
    message_received = Signal(dict)   # application message from child
    child_hwnd_ready = Signal(int)    # HWND sent by child on startup

    def __init__(self, parent: QObject = None):
        super().__init__(parent)
        global _ipc_counter
        _ipc_counter += 1
        self._socket: QLocalSocket | None = None
        self._buf    = b""
        self._server = QLocalServer(self)
        self._server.newConnection.connect(self._on_connection)
        name = f"pyxll-claude-{os.getpid()}-{_ipc_counter}"
        if not self._server.listen(name):
            QLocalServer.removeServer(name)
            if not self._server.listen(name):
                _log.error("QLocalServer.listen(%s) failed: %s",
                           name, self._server.errorString())

    @property
    def server_name(self) -> str:
        return self._server.serverName()

    def send(self, msg: dict) -> None:
        if self._socket and self._socket.isOpen():
            self._socket.write((json.dumps(msg) + "\n").encode())

    def close(self) -> None:
        if self._socket:
            self._socket.close()
            self._socket = None
        self._server.close()

    def _on_connection(self) -> None:
        self._socket = self._server.nextPendingConnection()
        self._socket.readyRead.connect(self._on_data)

    def _on_data(self) -> None:
        self._buf += bytes(self._socket.readAll())
        while b"\n" in self._buf:
            line, self._buf = self._buf.split(b"\n", 1)
            line = line.strip()
            if line:
                msg = json.loads(line.decode())
                if msg.get("type") == "ready":
                    self.child_hwnd_ready.emit(msg["hwnd"])
                else:
                    self.message_received.emit(msg)


# ---------------------------------------------------------------------------
# Client widget
# ---------------------------------------------------------------------------

class WebViewClient(QWidget):
    """Parent-process widget that embeds a child-process QWebEngineView.

    The child process runs child_process.py with a single QWebEngineView loaded
    at *url*.  Once the child sends its HWND, this widget embeds that window and
    wires up Windows thread-input queues so keyboard events work across processes.

    Signals:
        message_received(dict)  — application message from the child JS bridge
        child_ready()           — child window successfully embedded
    """
    message_received = Signal(dict)
    child_ready      = Signal()

    def __init__(
        self,
        url: str,
        xl_hwnd: int = 0,
        loading_text: str = "Loading…",
        parent=None,
    ):
        super().__init__(parent)
        self._url          = url
        self._xl_hwnd      = xl_hwnd
        self._child_hwnd   = 0
        self._child_container: QWidget | None = None
        self._is_destroyed = False

        self._ipc      = _IPC(self)
        self._child_qp = QProcess(self)

        self._loading_label = QLabel(loading_text, self)
        self._loading_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self._loading_label)

        self._ipc.child_hwnd_ready.connect(self._on_child_hwnd_ready)
        self._ipc.message_received.connect(self.message_received)
        self._child_qp.readyReadStandardError.connect(self._on_child_stderr)
        self._child_qp.finished.connect(self._on_child_finished)
        self._child_qp.errorOccurred.connect(self._on_child_error)
        self.destroyed.connect(self._cleanup)

        self._start_child()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def send_message(self, msg: dict) -> None:
        """Send a message to the child JS bridge."""
        self._ipc.send(msg)

    def focus_child(self) -> None:
        """Transfer keyboard focus to the child window."""
        if self._child_hwnd:
            _focus_foreign_hwnd(self._child_hwnd)

    def refocus_renderer(self) -> None:
        """Re-attach Chromium threads and focus Chrome_RenderWidgetHostHWND.

        After a mouse text selection Chromium moves Win32 focus to an internal
        HWND whose thread may not be in the attached set, breaking keyboard
        delivery to the page.  Re-running attach_child_threads picks up any new
        threads, then _focus_foreign_hwnd on the renderer HWND (re-attaching
        the caller's thread and calling SetFocus) restores key delivery.
        """
        if not self._child_hwnd or not self._xl_hwnd:
            return
        self.attach_child_threads()

        user32 = ctypes.windll.user32
        renderer_hwnd = [0]

        def _cb(hwnd: int, _lparam: int) -> bool:
            buf = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, buf, 256)
            if buf.value == "Chrome_RenderWidgetHostHWND":
                renderer_hwnd[0] = hwnd
                return False
            return True

        _EnumProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_ssize_t)
        user32.EnumChildWindows(self._child_hwnd, _EnumProc(_cb), 0)

        target = renderer_hwnd[0] or self._child_hwnd
        _log.debug("refocus_renderer: targeting hwnd=0x%x", target)
        _focus_foreign_hwnd(target)

    def showEvent(self, event) -> None:
        super().showEvent(event)
        if self._child_hwnd:
            _focus_foreign_hwnd(self._child_hwnd)

    def attach_child_threads(self) -> None:
        """Attach every Chromium child thread to Excel's input queue.

        Chromium creates internal HWNDs (compositor, render widget, etc.) whose
        threads are not attached automatically.  Call this once after the page's
        JS confirms Chromium's renderer is live (e.g. on 'start_process' or
        'editor_ready') so that Chromium's own internal SetFocus calls succeed
        and WM_KEYDOWN is delivered to JS.
        """
        if not self._child_hwnd or not self._xl_hwnd:
            return
        user32 = ctypes.windll.user32
        xl_tid = user32.GetWindowThreadProcessId(self._xl_hwnd, None)
        if not xl_tid:
            return
        seen: set[int] = set()

        def _cb(hwnd: int, _lparam: int) -> bool:
            tid = user32.GetWindowThreadProcessId(hwnd, None)
            if tid and tid not in seen:
                seen.add(tid)
                if tid != xl_tid:
                    ok = user32.AttachThreadInput(tid, xl_tid, True)
                    buf = ctypes.create_unicode_buffer(256)
                    user32.GetClassNameW(hwnd, buf, 256)
                    _log.debug(
                        "attach_child_threads: hwnd=0x%x class=%r tid=%d → %d",
                        hwnd, buf.value, tid, ok,
                    )
            return True

        _EnumProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_ssize_t)
        user32.EnumChildWindows(self._child_hwnd, _EnumProc(_cb), 0)

    # ------------------------------------------------------------------
    # Child process management
    # ------------------------------------------------------------------

    def _start_child(self) -> None:
        self._child_qp.setProgram(sys.executable)
        self._child_qp.setArguments([
            str(_CHILD_SCRIPT),
            "--ipc-name", self._ipc.server_name,
            "--url",      self._url,
        ])
        self._child_qp.start()

    def _on_child_error(self, error: QProcess.ProcessError) -> None:
        if not self._is_destroyed:
            _log.error("WebView child process error: %s", error)

    def _on_child_stderr(self) -> None:
        data = bytes(self._child_qp.readAllStandardError())
        for line in data.decode("utf-8", errors="replace").splitlines():
            if line.strip():
                _log.warning("[child] %s", line)

    def _on_child_finished(self, exit_code: int, _exit_status) -> None:
        if exit_code != 0 and not self._is_destroyed:
            _log.warning("WebView child process exited (code=%d)", exit_code)
        else:
            _log.debug("WebView child process exited (code=%d)", exit_code)

    # ------------------------------------------------------------------
    # Window embedding
    # ------------------------------------------------------------------

    @Slot(int)
    def _on_child_hwnd_ready(self, hwnd: int) -> None:
        """Embed the child window and wire up Windows input queues."""
        self._child_hwnd = hwnd
        self._loading_label.hide()

        foreign = QWindow.fromWinId(hwnd)
        container = QWidget.createWindowContainer(foreign)
        container.setMinimumSize(1, 1)
        container.installEventFilter(self)
        self.layout().addWidget(container)
        self._child_container = container

        user32   = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        my_tid   = kernel32.GetCurrentThreadId()
        child_tid = user32.GetWindowThreadProcessId(hwnd, None)

        # After SetParent (done by createWindowContainer), the child window
        # must have WS_CHILD style so Windows uses child-window activation rules.
        _GWL_STYLE   = -16
        _WS_CHILD    = 0x40000000
        _WS_POPUP    = 0x80000000
        _SWP_NOFLAGS = 0x0027   # NOMOVE | NOSIZE | NOZORDER | FRAMECHANGED
        style     = user32.GetWindowLongW(hwnd, _GWL_STYLE)
        new_style = (style & ~_WS_POPUP) | _WS_CHILD
        user32.SetWindowLongW(hwnd, _GWL_STYLE, new_style)
        user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, _SWP_NOFLAGS)

        # Attach thread input queues so SetFocus works across processes.
        xl_hwnd = self._xl_hwnd
        xl_tid  = user32.GetWindowThreadProcessId(xl_hwnd, None) if xl_hwnd else 0
        for (a, b, label) in [
            (child_tid, xl_tid,   "child→excel"),
            (my_tid,    xl_tid,   "parent→excel"),
            (my_tid,    child_tid, "parent→child"),
        ]:
            if a and b and a != b:
                ok = user32.AttachThreadInput(a, b, True)
                if ok != 1:
                    _log.warning("AttachThreadInput(%d, %d, True) [%s] → %d",
                                 a, b, label, ok)

        _focus_foreign_hwnd(hwnd)
        self.child_ready.emit()

    def eventFilter(self, obj, event: QEvent) -> bool:
        if obj is self._child_container:
            t = event.type()
            # Block cross-process WM_KILLFOCUS teardowns triggered by Qt's
            # QWindowContainer when clicks land on the embedded Chromium view.
            if t in (QEvent.Type.WindowActivate, QEvent.Type.WindowDeactivate):
                return True
            if self._child_hwnd and t == QEvent.Type.FocusIn:
                _log.debug("eventFilter: FocusIn → forwarding to child HWND")
                _focus_foreign_hwnd(self._child_hwnd)
        return super().eventFilter(obj, event)

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------

    def _cleanup(self, _obj=None) -> None:
        if self._is_destroyed:
            return
        self._is_destroyed = True

        if hasattr(self, "_ipc") and self._ipc:
            self._ipc.close()
            QApplication.processEvents()

        if hasattr(self, "_child_qp") and self._child_qp:
            if self._child_qp.state() != QProcess.ProcessState.NotRunning:
                if not self._child_qp.waitForFinished(1000):
                    self._child_qp.terminate()
                    if not self._child_qp.waitForFinished(5000):
                        _log.warning("WebView child did not terminate gracefully; killing")
                        self._child_qp.kill()
            self._child_qp = None

        # Detach keyboard input queues.
        user32   = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        my_tid   = kernel32.GetCurrentThreadId()
        xl_tid   = (user32.GetWindowThreadProcessId(self._xl_hwnd, None)
                    if self._xl_hwnd else 0)
        child_tid = (user32.GetWindowThreadProcessId(self._child_hwnd, None)
                     if self._child_hwnd else 0)
        for (a, b) in [(child_tid, xl_tid), (my_tid, xl_tid), (my_tid, child_tid)]:
            if a and b and a != b:
                try:
                    user32.AttachThreadInput(a, b, False)
                except Exception:
                    pass

    def closeEvent(self, event) -> None:
        self._cleanup()
        super().closeEvent(event)
