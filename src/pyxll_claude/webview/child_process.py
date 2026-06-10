"""
Child process hosting a single QWebEngineView.

Launched by WebViewClient:
    sys.executable path/to/child_process.py --ipc-name <name> --url <url>

Connects to the parent's QLocalServer, sends its HWND, then relays messages
between the JS bridge and the parent process via newline-delimited JSON IPC.

Intentionally imports nothing from pyxll_claude so it can start cleanly as a
plain subprocess without Excel or the PyXLL package being available.
"""
import argparse
import ctypes
import ctypes.wintypes
import json
import os
import sys

from PySide6.QtCore import (
    QAbstractNativeEventFilter, QEvent, QObject, Qt, QTimer, Signal, Slot,
)
from PySide6.QtNetwork import QLocalSocket
from PySide6.QtWebChannel import QWebChannel
from PySide6.QtWebEngineCore import QWebEnginePage, QWebEngineSettings
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWidgets import QApplication, QVBoxLayout, QWidget


# ---------------------------------------------------------------------------
# Native WM_SETFOCUS filter
# ---------------------------------------------------------------------------

class _MSG_c(ctypes.Structure):
    _fields_ = [
        ("hwnd",    ctypes.c_void_p),
        ("message", ctypes.c_uint),
        ("wParam",  ctypes.c_size_t),
        ("lParam",  ctypes.c_ssize_t),
        ("time",    ctypes.c_ulong),
        ("pt_x",    ctypes.c_long),
        ("pt_y",    ctypes.c_long),
    ]


class _SetFocusFilter(QAbstractNativeEventFilter):
    """Forward Qt focus to the view whenever the outer HWND gets WM_SETFOCUS.

    After the parent changes the child window style to WS_CHILD, Windows no
    longer sends WM_ACTIVATE, so QEvent::WindowActivate never fires.  We catch
    WM_SETFOCUS directly so that QApplication::focusWidget() is always the view,
    which is necessary for WM_KEYDOWN messages to reach JS.
    """
    _WM_SETFOCUS    = 0x0007
    _WM_LBUTTONDOWN = 0x0201
    _WM_RBUTTONDOWN = 0x0204
    _WM_MBUTTONDOWN = 0x0207

    def __init__(self, win_hwnd: int, get_view):
        super().__init__()
        self._win_hwnd  = win_hwnd
        self._get_view  = get_view

    def nativeEventFilter(self, event_type, message):
        if event_type != b"windows_generic_MSG":
            return False, 0
        msg  = ctypes.cast(int(message), ctypes.POINTER(_MSG_c)).contents
        hwnd = msg.hwnd or 0
        t    = msg.message

        if t in (self._WM_LBUTTONDOWN, self._WM_RBUTTONDOWN, self._WM_MBUTTONDOWN):
            user32 = ctypes.windll.user32
            if hwnd == self._win_hwnd or user32.IsChild(self._win_hwnd, hwnd):
                cur = user32.GetFocus()
                if cur == self._win_hwnd or user32.IsChild(self._win_hwnd, cur):
                    return False, 0
                user32.SetFocus(self._win_hwnd)

        if t == self._WM_SETFOCUS and hwnd == self._win_hwnd:
            view = self._get_view()
            if view:
                view.setFocus(Qt.FocusReason.ActiveWindowFocusReason)

        return False, 0


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class _MainWindow(QWidget):
    """Outer frameless window for the child process."""

    def __init__(self):
        super().__init__()
        self._view: QWebEngineView | None = None

    def set_view(self, view: QWebEngineView) -> None:
        self._view = view

    def focusInEvent(self, event: QEvent) -> None:
        if self._view:
            self._view.setFocus(Qt.FocusReason.OtherFocusReason)
        super().focusInEvent(event)

    def changeEvent(self, event: QEvent) -> None:
        super().changeEvent(event)
        # Fallback for the top-level window case (before WS_CHILD is applied).
        if event.type() == QEvent.Type.WindowActivate and self._view:
            self._view.setFocus(Qt.FocusReason.OtherFocusReason)


# ---------------------------------------------------------------------------
# IPC — child side (socket client)
# ---------------------------------------------------------------------------

class _IPC(QObject):
    """QLocalSocket connection to the parent, framed as newline-delimited JSON.

    Incoming messages (parent → child) are emitted as raw JSON strings via
    message_received so the bridge can forward them directly to JS.
    """
    message_received = Signal(str)  # raw JSON string from parent

    def __init__(self, name: str, parent: QObject = None):
        super().__init__(parent)
        self._name    = name
        self._buf     = b""
        self._retries = 0
        self._hwnd    = 0
        self._sock    = QLocalSocket(self)
        self._sock.connected.connect(self._on_connected)
        self._sock.readyRead.connect(self._on_data)
        self._sock.disconnected.connect(self._on_disconnected)
        self._sock.errorOccurred.connect(self._on_error)
        self._sock.connectToServer(name)

    def set_hwnd(self, hwnd: int) -> None:
        """Record HWND; send 'ready' immediately if already connected."""
        self._hwnd = hwnd
        if self._sock.state() == QLocalSocket.LocalSocketState.ConnectedState:
            self._send_ready()

    def send(self, msg: dict) -> None:
        self._sock.write((json.dumps(msg) + "\n").encode())

    def _on_connected(self) -> None:
        if self._hwnd:
            self._send_ready()

    def _send_ready(self) -> None:
        self.send({"type": "ready", "hwnd": self._hwnd})

    def _on_disconnected(self) -> None:
        # Parent closed the connection.  QWebEngine's shutdown sequence blocks
        # the Qt event loop, so QApplication.quit() + a timer fallback never
        # fires reliably.  The child process has no state to save, so just exit.
        os._exit(0)

    def _on_error(self, _error) -> None:
        if (self._sock.state() != QLocalSocket.LocalSocketState.ConnectedState
                and self._retries < 10):
            self._retries += 1
            QTimer.singleShot(500, lambda: self._sock.connectToServer(self._name))

    def _on_data(self) -> None:
        self._buf += bytes(self._sock.readAll())
        while b"\n" in self._buf:
            line, self._buf = self._buf.split(b"\n", 1)
            line = line.strip()
            if line:
                self.message_received.emit(line.decode())


# ---------------------------------------------------------------------------
# General-purpose QWebChannel bridge
# ---------------------------------------------------------------------------

class _Bridge(QObject):
    """Single bridge class registered as 'bridge' in every web view.

    JS → Python: bridge.postMessage(jsonString)
    Python → JS: bridge.messageReceived signal (jsonString)

    Clipboard helpers are exposed directly for convenience since they return a
    value and don't fit the fire-and-forget postMessage pattern.
    """
    messageReceived = Signal(str)   # JSON string pushed to JS

    def __init__(self, ipc: "_IPC", view: QWebEngineView, parent: QObject = None):
        super().__init__(parent)
        self._ipc  = ipc
        self._view = view
        ipc.message_received.connect(self.messageReceived)

    @Slot(str)
    def postMessage(self, data: str) -> None:
        """Called from JS to send a message to the parent process."""
        self._ipc.send(json.loads(data))

    @Slot(str)
    def copyToClipboard(self, text: str) -> None:
        QApplication.clipboard().setText(text)

    @Slot(result=str)
    def pasteFromClipboard(self) -> str:
        return QApplication.clipboard().text()


# ---------------------------------------------------------------------------
# Logged web page
# ---------------------------------------------------------------------------

class _LoggedPage(QWebEnginePage):
    def javaScriptConsoleMessage(self, level, message, line, source_id):
        if level == QWebEnginePage.JavaScriptConsoleMessageLevel.ErrorMessageLevel:
            print(f"[child] JS ERROR: {message} ({source_id}:{line})",
                  file=sys.stderr, flush=True)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--ipc-name", required=True)
    parser.add_argument("--url",      required=True)
    args = parser.parse_args()

    app = QApplication(sys.argv)

    win = _MainWindow()
    win.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool)

    ipc = _IPC(args.ipc_name, win)

    view = QWebEngineView(win)
    view.setPage(_LoggedPage(view))
    view.page().settings().setAttribute(
        QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls, True
    )
    view.page().settings().setAttribute(
        QWebEngineSettings.WebAttribute.JavascriptCanAccessClipboard, True
    )

    channel = QWebChannel(view)
    bridge  = _Bridge(ipc, view, view)
    channel.registerObject("bridge", bridge)
    view.page().setWebChannel(channel)

    from PySide6.QtCore import QUrl
    url = args.url
    if os.path.isabs(url):
        view.setUrl(QUrl.fromLocalFile(url))
    else:
        view.setUrl(QUrl(url))

    root = QVBoxLayout(win)
    root.setContentsMargins(0, 0, 0, 0)
    root.addWidget(view)
    win.set_view(view)

    win.resize(800, 600)
    win_hwnd = int(win.winId())
    ipc.set_hwnd(win_hwnd)
    win.show()

    # Give the view explicit Qt focus so WM_KEYDOWN messages reach JS.
    view.setFocus(Qt.FocusReason.OtherFocusReason)

    focus_filter = _SetFocusFilter(win_hwnd, lambda: view)
    app.installNativeEventFilter(focus_filter)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
