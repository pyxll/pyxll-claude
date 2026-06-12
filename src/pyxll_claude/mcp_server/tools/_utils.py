"""Shared helpers for MCP tool implementations."""
from __future__ import annotations

import threading
from typing import Any

import pyxll


def run_on_main_thread(func, timeout: float = 15.0) -> Any:
    """Run func() on Excel's main thread and return its result.

    Raises TimeoutError if Excel doesn't respond, or re-raises whatever func raised.
    """
    result: list[Any] = []
    error: list[BaseException] = []
    done = threading.Event()

    def _wrapper():
        try:
            result.append(func())
        except BaseException as exc:  # noqa: BLE001
            error.append(exc)
        finally:
            done.set()

    pyxll.schedule_call(_wrapper)
    if not done.wait(timeout=timeout):
        raise TimeoutError("Excel did not respond (it may be busy)")
    if error:
        raise error[0]
    return result[0] if result else None


def get_sheet(xl, name: str | None):
    """Return the named sheet, or the active sheet if name is None."""
    return xl.Sheets(name) if name else xl.ActiveSheet


def hex_to_bgr(hex_color: str) -> int:
    """Convert '#RRGGBB' to the BGR integer Excel's Color property expects."""
    hex_color = hex_color.lstrip("#")
    if len(hex_color) != 6:
        raise ValueError(f"Expected a 6-digit hex color like '#RRGGBB', got {hex_color!r}")
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return b + (g * 256) + (r * 65536)


def bgr_to_hex(color: int) -> str:
    """Convert an Excel BGR integer back to '#RRGGBB'."""
    b = color % 256
    g = (color // 256) % 256
    r = (color // 65536) % 256
    return f"#{r:02x}{g:02x}{b:02x}"


def to_json_safe(v: Any) -> Any:
    """Convert a COM value to something json.dumps can handle."""
    if isinstance(v, (type(None), bool, int, float, str)):
        return v
    return str(v)
