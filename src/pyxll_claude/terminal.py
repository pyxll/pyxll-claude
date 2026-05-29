"""
Windows ConPTY wrapper for the claude CLI subprocess.

All PTY operations (create, read, write, resize, destroy) run on a single
daemon thread — _pty_loop — so pywinpty is never called from more than one
thread at a time.  Communication works in two directions:

  main thread → PTY thread : _cmd_queue  (_WriteCmd, _ResizeCmd, _STOP)
  PTY thread → main thread : QCoreApplication.postEvent  (_DataEvent, _FinishedEvent)

QCoreApplication.postEvent is thread-safe (C++ mutex-guarded) and delivers
events to ClaudeProcess.event() on the Qt main thread, avoiding the
Python-level locking that caused the concurrent.futures race under Python 3.14.
"""
import logging
import queue
import shutil
import threading
from pathlib import Path

from PySide6.QtCore import QCoreApplication, QEvent, QObject, Signal

_log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Command types sent from the main thread to the PTY thread
# ---------------------------------------------------------------------------

_STOP = object()  # sentinel — tells the PTY thread to exit


class _WriteCmd:
    __slots__ = ("data",)
    def __init__(self, data: bytes) -> None:
        self.data = data


class _ResizeCmd:
    __slots__ = ("cols", "rows")
    def __init__(self, cols: int, rows: int) -> None:
        self.cols = cols
        self.rows = rows


# ---------------------------------------------------------------------------
# Qt events sent from the PTY thread back to the main thread
# ---------------------------------------------------------------------------

class _DataEvent(QEvent):
    _type: QEvent.Type = QEvent.Type(QEvent.registerEventType())

    def __init__(self, data: bytes) -> None:
        super().__init__(self._type)
        self.data = data


class _FinishedEvent(QEvent):
    _type: QEvent.Type = QEvent.Type(QEvent.registerEventType())

    def __init__(self) -> None:
        super().__init__(self._type)


# ---------------------------------------------------------------------------
# Project-root fallback (used when no workspace is configured)
# ---------------------------------------------------------------------------

def _find_project_root() -> Path:
    current = Path(__file__).resolve()
    for _ in range(8):
        current = current.parent
        if (current / "CLAUDE.md").exists():
            return current
    return Path(__file__).resolve().parent.parent.parent


# ---------------------------------------------------------------------------
# Main class
# ---------------------------------------------------------------------------

class ClaudeProcess(QObject):
    """Runs the claude CLI inside a Windows ConPTY.

    All PTY interaction is confined to a single daemon thread (_pty_loop).
    The main thread sends commands via a queue and receives output as Qt events.
    """

    data_received = Signal(bytes)
    finished = Signal()

    def __init__(self, parent: QObject = None) -> None:
        super().__init__(parent)
        self._cmd_queue: queue.Queue = queue.Queue()
        self._pty_thread: threading.Thread | None = None
        self._running = False

    # ------------------------------------------------------------------
    # Public API — called from the Qt main thread
    # ------------------------------------------------------------------

    def start(self, cols: int = 80, rows: int = 24, workspace: Path = None) -> None:
        """Spawn claude in a ConPTY.  All PTY work happens on the PTY thread."""
        if workspace is None:
            workspace = _find_project_root()
        _log.info("Starting claude CLI in %s (cols=%d rows=%d)", workspace, cols, rows)
        self._running = True
        self._pty_thread = threading.Thread(
            target=self._pty_loop,
            args=(workspace, cols, rows),
            daemon=True,
        )
        self._pty_thread.start()

    def write(self, data: bytes) -> None:
        """Enqueue bytes to be written to the PTY stdin."""
        if self._running:
            self._cmd_queue.put(_WriteCmd(data))

    def resize(self, cols: int, rows: int) -> None:
        """Enqueue a terminal size change for the PTY."""
        if self._running:
            self._cmd_queue.put(_ResizeCmd(cols, rows))

    def terminate(self) -> None:
        """Ask the PTY thread to stop."""
        self._running = False
        self._cmd_queue.put(_STOP)

    # ------------------------------------------------------------------
    # Qt event handler — receives events posted by the PTY thread
    # ------------------------------------------------------------------

    def event(self, event: QEvent) -> bool:
        if isinstance(event, _DataEvent):
            self.data_received.emit(event.data)
            return True
        if isinstance(event, _FinishedEvent):
            self.finished.emit()
            return True
        return super().event(event)

    # ------------------------------------------------------------------
    # PTY thread
    # ------------------------------------------------------------------

    def _handle_cmd(self, cmd: object, pty: object) -> bool:
        """Execute one command on the PTY thread.  Returns True to stop."""
        if cmd is _STOP:
            return True
        if isinstance(cmd, _WriteCmd):
            try:
                pty.write(cmd.data.decode("utf-8", errors="replace"))
            except Exception as exc:
                _log.error("PTY write error: %s", exc)
                return True
        elif isinstance(cmd, _ResizeCmd):
            try:
                pty.set_size(cmd.cols, cmd.rows)
            except Exception as exc:
                _log.warning("PTY resize failed: %s", exc)
        return False

    def _pty_loop(self, workspace: Path, cols: int, rows: int) -> None:
        """Daemon thread: creates, operates, and destroys the PTY.

        Loop behaviour:
          1. Non-blocking read — post any available output and loop immediately.
          2. If no output and process is alive, block on the command queue for
             up to 20 ms so user input and resize events are handled promptly.
          3. Drain all remaining queued commands before going back to read.
        """
        from winpty import PTY

        claude = shutil.which("claude")
        if claude is None:
            _log.error("claude CLI not found on PATH")
            QCoreApplication.postEvent(self, _FinishedEvent())
            return

        appname = (
            f'cmd.exe /d /c "{claude}"'
            if Path(claude).suffix.lower() in (".cmd", ".bat")
            else claude
        )

        pty = PTY(cols, rows)
        try:
            pty.spawn(appname, cwd=str(workspace))

            stop = False
            while not stop:
                # 1. Non-blocking read — drain all available output first.
                try:
                    data = pty.read(blocking=False)
                except Exception as exc:
                    _log.error("PTY read error: %s", exc)
                    break

                if data:
                    if isinstance(data, str):
                        data = data.encode("utf-8", errors="replace")
                    QCoreApplication.postEvent(self, _DataEvent(data))
                    continue  # loop back immediately to read more

                # 2. No output — check whether the process is still running.
                if not pty.isalive():
                    break

                # 3. Block briefly on the command queue so we respond quickly
                #    to user keystrokes and resize events.
                try:
                    cmd = self._cmd_queue.get(timeout=0.02)
                except queue.Empty:
                    continue  # nothing pending — loop back to read

                stop = self._handle_cmd(cmd, pty)

                # 4. Drain any additional commands that arrived while we were
                #    blocked, so they don't pile up behind the next read.
                while not stop:
                    try:
                        cmd = self._cmd_queue.get_nowait()
                    except queue.Empty:
                        break
                    stop = self._handle_cmd(cmd, pty)

        finally:
            try:
                del pty
            except Exception as exc:
                _log.warning("Error closing PTY: %s", exc)

        self._running = False
        QCoreApplication.postEvent(self, _FinishedEvent())
