import json
import traceback

from pyxll import xl_app

from . import register
from ._utils import bgr_to_hex, get_sheet, run_on_main_thread, to_json_safe

_DEFAULT_WHITE = 16777215  # Excel Interior.Color when no fill
_DEFAULT_BLACK = 0         # Excel Font.Color when default black


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    ref = args.get("ref")
    sheet = args.get("sheet")
    try:
        def _get():
            xl = xl_app()
            rng = get_sheet(xl, sheet).Range(ref)

            fill_color = None
            try:
                c = rng.Interior.Color
                if c is not None and int(c) != _DEFAULT_WHITE:
                    fill_color = bgr_to_hex(int(c))
            except Exception:
                pass

            font_color = None
            try:
                c = rng.Font.Color
                if c is not None and int(c) != _DEFAULT_BLACK:
                    font_color = bgr_to_hex(int(c))
            except Exception:
                pass

            return {
                "value": to_json_safe(rng.Value),
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

        return json.dumps(run_on_main_thread(_get)), False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


register(
    "pyxll_get_cell_info",
    (
        "Return detailed information about a single cell as a JSON object: value, formula, "
        "number format, font (bold, italic, name, size, color), fill color, column width, "
        "and row height. Use this when you need to inspect styling or formatting, not just "
        "the raw value — prefer pyxll_read_range for bulk value reads."
    ),
    {
        "type": "object",
        "properties": {
            "ref": {
                "type": "string",
                "description": "A1 reference to a single cell, e.g. 'B3'.",
            },
            "sheet": {
                "type": "string",
                "description": "Worksheet name. Omit to use the active sheet.",
            },
        },
        "required": ["ref"],
    },
    _impl,
)
