"""Excel COM operations for the PyXLL-Claude MCP server.

Excel's COM API is only safe to call from Excel's main thread. Every operation
here is marshalled there via ``pyxll.schedule_call`` and blocks until it returns,
using the same ``schedule_call`` + ``threading.Event`` pattern as the inline
tools in ``mcp_server.py``. Functions return plain Python values; the MCP layer
is responsible for JSON-encoding them.
"""
import base64
import io
import threading
import time
from typing import Any

import pyxll
from pyxll import xl_app

# Upper bound on bulk row/column inserts. Each insert is a separate COM call on
# Excel's main thread, and run_on_main_thread's timeout unblocks the waiting
# thread but does NOT cancel a callback already running — so an unbounded count
# would hang Excel. Reject oversized requests up front instead.
_MAX_INSERT = 1000


def run_on_main_thread(func, timeout: float = 15.0):
    """Run ``func()`` on Excel's main thread and return its result.

    Args:
        func: Zero-argument callable performing COM operations.
        timeout: Seconds to wait for Excel to service the call.

    Raises:
        TimeoutError: If Excel does not respond within ``timeout``.
        Exception: Whatever ``func`` raised, re-raised on the calling thread.
    """
    result: list[Any] = []
    error: list[BaseException] = []
    done = threading.Event()

    def _wrapper():
        try:
            result.append(func())
        except BaseException as exc:  # noqa: BLE001 - re-raised below
            error.append(exc)
        finally:
            done.set()

    pyxll.schedule_call(_wrapper)
    if not done.wait(timeout=timeout):
        raise TimeoutError("Excel did not respond (it may be busy)")
    if error:
        raise error[0]
    return result[0] if result else None


def _sheet(xl, sheet: str | None):
    """Resolve a worksheet by name, defaulting to the active sheet."""
    return xl.Sheets(sheet) if sheet else xl.ActiveSheet


def _hex_to_rgb_int(hex_color: str) -> int:
    """Convert ``#RRGGBB`` to the BGR integer Excel's ``Color`` property expects."""
    hex_color = hex_color.lstrip("#")
    if len(hex_color) != 6:
        raise ValueError(f"Expected a 6-digit hex color like '#RRGGBB', got {hex_color!r}")
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return b + (g * 256) + (r * 256 * 256)


def _bgr_int_to_hex(color: int) -> str:
    """Convert an Excel BGR integer back to ``#RRGGBB``."""
    b = color % 256
    g = (color // 256) % 256
    r = (color // 65536) % 256
    return f"#{r:02x}{g:02x}{b:02x}"


# ---------------------------------------------------------------------------
# Reading / inspection
# ---------------------------------------------------------------------------

def get_sheets() -> list[str]:
    """Return the names of every sheet in the active workbook."""
    def _get():
        xl = xl_app()
        if xl.ActiveWorkbook is None:
            return []
        return [s.Name for s in xl.ActiveWorkbook.Sheets]

    return run_on_main_thread(_get)


def get_used_range(sheet: str | None = None) -> dict[str, Any]:
    """Return the bounds of the populated area of a sheet."""
    def _get():
        xl = xl_app()
        used = _sheet(xl, sheet).UsedRange
        return {
            "address": used.Address.replace("$", ""),
            "row_count": used.Rows.Count,
            "column_count": used.Columns.Count,
            "start_row": used.Row,
            "start_column": used.Column,
        }

    return run_on_main_thread(_get)


def get_cell_info(ref: str, sheet: str | None = None) -> dict[str, Any]:
    """Return value, formula, number format and basic styling for one cell."""
    def _get():
        xl = xl_app()
        rng = _sheet(xl, sheet).Range(ref)

        fill_color = None
        try:
            interior = rng.Interior.Color
            if interior is not None and interior != 16777215:  # not the default white
                fill_color = _bgr_int_to_hex(int(interior))
        except Exception:
            pass

        font_color = None
        try:
            fc = rng.Font.Color
            if fc is not None and fc != 0:  # not the default black
                font_color = _bgr_int_to_hex(int(fc))
        except Exception:
            pass

        return {
            "value": rng.Value,
            "formula": rng.Formula2 or "",
            "number_format": rng.NumberFormat,
            "font_bold": bool(rng.Font.Bold),
            "font_italic": bool(rng.Font.Italic),
            "font_name": rng.Font.Name,
            "font_size": rng.Font.Size,
            "font_color": font_color,
            "fill_color": fill_color,
            "column_width": rng.ColumnWidth,
            "row_height": rng.RowHeight,
        }

    return run_on_main_thread(_get)


def get_named_ranges() -> list[dict[str, str]]:
    """Return every defined name in the active workbook."""
    def _get():
        xl = xl_app()
        wb = xl.ActiveWorkbook
        if wb is None:
            return []
        return [{"name": n.Name, "refers_to": n.RefersTo} for n in wb.Names]

    return run_on_main_thread(_get)


# ---------------------------------------------------------------------------
# Writing single cells
# ---------------------------------------------------------------------------

def write_value(ref: str, value: Any, sheet: str | None = None) -> None:
    """Write a single value to a cell."""
    def _write():
        xl = xl_app()
        _sheet(xl, sheet).Range(ref).Value2 = value

    run_on_main_thread(_write)


def write_formula(ref: str, formula: str, sheet: str | None = None) -> None:
    """Write a formula to a cell (e.g. ``=SUM(A1:A10)``)."""
    def _write():
        xl = xl_app()
        _sheet(xl, sheet).Range(ref).Formula2 = formula

    run_on_main_thread(_write)


# ---------------------------------------------------------------------------
# Structure
# ---------------------------------------------------------------------------

def clear_range(ref: str, clear_type: str = "all", sheet: str | None = None) -> None:
    """Clear ``contents``, ``formats`` or ``all`` from a range (no row/col deletion)."""
    def _clear():
        xl = xl_app()
        rng = _sheet(xl, sheet).Range(ref)
        if clear_type == "contents":
            rng.ClearContents()
        elif clear_type == "formats":
            rng.ClearFormats()
        else:
            rng.Clear()

    run_on_main_thread(_clear)


def insert_rows(row: int, count: int = 1, sheet: str | None = None) -> None:
    """Insert ``count`` rows above ``row`` (1-indexed)."""
    if not 1 <= count <= _MAX_INSERT:
        raise ValueError(f"count must be between 1 and {_MAX_INSERT}, got {count}")

    def _insert():
        xl = xl_app()
        ws = _sheet(xl, sheet)
        for _ in range(count):
            ws.Rows(row).Insert()

    run_on_main_thread(_insert)


def insert_columns(column: str, count: int = 1, sheet: str | None = None) -> None:
    """Insert ``count`` columns before ``column`` (e.g. ``'B'``)."""
    if not 1 <= count <= _MAX_INSERT:
        raise ValueError(f"count must be between 1 and {_MAX_INSERT}, got {count}")

    def _insert():
        xl = xl_app()
        ws = _sheet(xl, sheet)
        for _ in range(count):
            ws.Columns(column).Insert()

    run_on_main_thread(_insert)


def merge_cells(ref: str, sheet: str | None = None) -> None:
    """Merge the cells of a range into one."""
    def _merge():
        xl = xl_app()
        _sheet(xl, sheet).Range(ref).Merge()

    run_on_main_thread(_merge)


def unmerge_cells(ref: str, sheet: str | None = None) -> None:
    """Unmerge previously merged cells."""
    def _unmerge():
        xl = xl_app()
        _sheet(xl, sheet).Range(ref).UnMerge()

    run_on_main_thread(_unmerge)


def add_sheet(name: str | None = None, after: str | None = None) -> str:
    """Add a worksheet to the active workbook and return its name."""
    def _add():
        xl = xl_app()
        wb = xl.ActiveWorkbook
        if wb is None:
            raise ValueError("No active workbook")
        anchor = wb.Sheets[after] if after else wb.Sheets(wb.Sheets.Count)
        new_sheet = wb.Sheets.Add(After=anchor)
        if name:
            new_sheet.Name = name
        return new_sheet.Name

    return run_on_main_thread(_add)


def name_range(name: str, ref: str, sheet: str | None = None) -> None:
    """Create or update a workbook-level named range."""
    def _name():
        xl = xl_app()
        wb = xl.ActiveWorkbook
        sheet_name = sheet or xl.ActiveSheet.Name
        refers_to = f"='{sheet_name}'!{ref}"
        try:
            wb.Names(name).RefersTo = refers_to
        except Exception:
            wb.Names.Add(Name=name, RefersTo=refers_to)

    run_on_main_thread(_name)


def save_workbook(path: str | None = None) -> str:
    """Save the active workbook (to ``path`` if given) and return its full path."""
    def _save():
        xl = xl_app()
        wb = xl.ActiveWorkbook
        if wb is None:
            raise ValueError("No active workbook")
        if path:
            wb.SaveAs(path)
        elif wb.Path:
            wb.Save()
        else:
            raise ValueError("Workbook has never been saved; provide a path.")
        return wb.FullName

    return run_on_main_thread(_save)


# ---------------------------------------------------------------------------
# Formatting
# ---------------------------------------------------------------------------

def format_cells(ref: str, options: dict[str, Any], sheet: str | None = None) -> None:
    """Apply font / fill / border / alignment / number-format options to a range.

    Recognised ``options`` keys: ``bold``, ``italic``, ``underline``,
    ``font_name``, ``font_size``, ``font_color``, ``fill_color``,
    ``number_format``, ``horizontal_alignment``, ``vertical_alignment``,
    ``wrap_text``, ``border``. Unknown keys are ignored.
    """
    def _format():
        xl = xl_app()
        rng = _sheet(xl, sheet).Range(ref)

        if options.get("bold") is not None:
            rng.Font.Bold = options["bold"]
        if options.get("italic") is not None:
            rng.Font.Italic = options["italic"]
        if options.get("underline") is not None:
            # xlUnderlineStyleSingle = 2, xlUnderlineStyleNone = -4142
            rng.Font.Underline = 2 if options["underline"] else -4142
        if options.get("font_name"):
            rng.Font.Name = options["font_name"]
        if options.get("font_size"):
            rng.Font.Size = options["font_size"]
        if options.get("font_color"):
            rng.Font.Color = _hex_to_rgb_int(options["font_color"])
        if options.get("fill_color"):
            rng.Interior.Color = _hex_to_rgb_int(options["fill_color"])
        if options.get("number_format"):
            rng.NumberFormat = options["number_format"]
        if options.get("horizontal_alignment"):
            h_map = {"left": -4131, "center": -4108, "right": -4152}
            rng.HorizontalAlignment = h_map.get(
                options["horizontal_alignment"].lower(), -4131
            )
        if options.get("vertical_alignment"):
            v_map = {"top": -4160, "center": -4108, "bottom": -4107}
            rng.VerticalAlignment = v_map.get(
                options["vertical_alignment"].lower(), -4108
            )
        if options.get("wrap_text") is not None:
            rng.WrapText = options["wrap_text"]
        if options.get("border"):
            # xlEdgeLeft(7) through xlInsideHorizontal(12); xlContinuous=1, xlThin=2
            for edge in range(7, 13):
                try:
                    border = rng.Borders(edge)
                    border.LineStyle = 1
                    border.Weight = 2
                except Exception:
                    pass  # some edges don't apply to every range shape

    run_on_main_thread(_format)


def auto_fit_column(column: str, sheet: str | None = None) -> None:
    """Auto-fit one or more columns to their content (e.g. ``'A'`` or ``'A:C'``)."""
    def _autofit():
        xl = xl_app()
        col_range = column if ":" in column else f"{column}:{column}"
        _sheet(xl, sheet).Range(col_range).Columns.AutoFit()

    run_on_main_thread(_autofit)


def set_column_width(column: str, width: float, sheet: str | None = None) -> None:
    """Set the width of a column or column range, in character units."""
    def _set():
        xl = xl_app()
        col_range = column if ":" in column else f"{column}:{column}"
        _sheet(xl, sheet).Range(col_range).ColumnWidth = width

    run_on_main_thread(_set)


def set_row_height(row: int, height: float, sheet: str | None = None) -> None:
    """Set the height of a row, in points."""
    def _set():
        xl = xl_app()
        _sheet(xl, sheet).Rows(row).RowHeight = height

    run_on_main_thread(_set)


# ---------------------------------------------------------------------------
# Calculation
# ---------------------------------------------------------------------------

def calculate(scope: str = "workbook", sheet: str | None = None,
              ref: str | None = None) -> None:
    """Force a recalculation of the whole app, a sheet, or a single range."""
    def _calc():
        xl = xl_app()
        if scope == "workbook":
            xl.Calculate()
        elif scope == "sheet":
            _sheet(xl, sheet).Calculate()
        elif scope == "range":
            if not ref:
                raise ValueError("ref is required for scope='range'")
            _sheet(xl, sheet).Range(ref).Calculate()
        else:
            raise ValueError(f"Invalid scope: {scope!r}")

    run_on_main_thread(_calc)


# ---------------------------------------------------------------------------
# Screenshot
# ---------------------------------------------------------------------------

def screenshot(ref: str | None = None, sheet: str | None = None) -> str:
    """Capture a range (or the whole used range) as a base64-encoded PNG.

    Uses ``Range.CopyPicture`` to place a bitmap on the clipboard, then reads it
    back with Pillow. If ``ref`` is omitted the sheet's ``UsedRange`` is captured.
    """
    def _capture():
        from PIL import ImageGrab  # imported lazily so the dep is only needed here

        xl = xl_app()
        ws = _sheet(xl, sheet)
        rng = ws.Range(ref) if ref else ws.UsedRange

        # xlScreen = 1 (Appearance), xlBitmap = 2 (Format)
        rng.CopyPicture(Appearance=1, Format=2)

        # Poll briefly rather than sleeping a fixed interval: it both shortens the
        # common case and narrows the window in which another process could clobber
        # the clipboard between CopyPicture and the read.
        img = None
        for _ in range(20):
            img = ImageGrab.grabclipboard()
            if img is not None:
                break
            time.sleep(0.05)
        if img is None:
            raise RuntimeError("Failed to read the captured image from the clipboard")

        buffer = io.BytesIO()
        img.save(buffer, format="PNG")
        return base64.b64encode(buffer.getvalue()).decode("ascii")

    return run_on_main_thread(_capture, timeout=30.0)
