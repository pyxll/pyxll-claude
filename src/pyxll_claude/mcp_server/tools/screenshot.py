import base64
import logging
import threading
import traceback
from typing import Any

import pyxll
from pyxll import xl_app

from . import register

_log = logging.getLogger(__name__)

# PrintWindow flag — renders GPU-accelerated content (requires Windows 8+).
_PW_RENDERFULLCONTENT = 2


def _capture_hwnd_to_png(hwnd: int) -> bytes:
    """Capture a native window by HWND and return PNG bytes via pywin32 + PySide6."""
    import win32gui
    import win32ui
    from ctypes import windll
    from PySide6.QtCore import QBuffer, QIODeviceBase
    from PySide6.QtGui import QImage

    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = right - left
    height = bottom - top

    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    mem_dc = mfc_dc.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(mfc_dc, width, height)
    mem_dc.SelectObject(bmp)

    try:
        windll.user32.PrintWindow(hwnd, mem_dc.GetSafeHdc(), _PW_RENDERFULLCONTENT)
        raw = bmp.GetBitmapBits(True)
    finally:
        mem_dc.DeleteDC()
        mfc_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwnd_dc)
        win32gui.DeleteObject(bmp.GetHandle())

    # GDI bitmaps are BGRA; QImage.Format_ARGB32 is BGRA in memory on little-endian.
    img = QImage(raw, width, height, QImage.Format.Format_ARGB32)
    buf = QBuffer()
    buf.open(QIODeviceBase.OpenModeFlag.WriteOnly)
    img.save(buf, "PNG")
    return bytes(buf.data())


def _impl(args: dict, ctx: dict) -> tuple[list[dict] | str, bool]:
    result: list[Any] = [None]
    done = threading.Event()

    def _do_screenshot():
        try:
            hwnd = xl_app().Hwnd
            png_bytes = _capture_hwnd_to_png(hwnd)
            b64 = base64.b64encode(png_bytes).decode("ascii")
            result[0] = [{"type": "image", "data": b64, "mimeType": "image/png"}]
        except Exception:
            result[0] = "ERROR: " + traceback.format_exc()
        finally:
            done.set()

    pyxll.schedule_call(_do_screenshot)
    done.wait(timeout=15)

    if result[0] is None:
        return "ERROR: screenshot timed out", True
    r = result[0]
    if isinstance(r, str):
        return r, True
    return r, False


register(
    "pyxll_screenshot",
    (
        "Capture a screenshot of the Excel application window and return it as an image. "
        "Use this to visually inspect the current state of the spreadsheet — e.g. to see "
        "charts, formatting, cell layout, or to verify that a write operation produced the "
        "expected result. Excel must be visible (not minimised) for a useful capture."
    ),
    {"type": "object", "properties": {}, "required": []},
    _impl,
)
