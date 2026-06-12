import traceback

from pyxll import xl_app

from . import register
from ._utils import get_sheet, hex_to_bgr, run_on_main_thread

# Excel xlUnderlineStyle constants
_XL_UNDERLINE_SINGLE = 2
_XL_UNDERLINE_NONE = -4142

# Excel horizontal alignment constants
_H_ALIGN = {"left": -4131, "center": -4108, "right": -4152}

# Excel vertical alignment constants
_V_ALIGN = {"top": -4160, "center": -4108, "bottom": -4107}

# xlEdgeLeft(7)..xlInsideHorizontal(12); xlContinuous=1, xlThin=2
_BORDER_EDGES = range(7, 13)


def _impl(args: dict, ctx: dict) -> tuple[str, bool]:
    ref = args.get("ref")
    options: dict = args.get("options", {})
    sheet = args.get("sheet")
    try:
        def _format():
            xl = xl_app()
            rng = get_sheet(xl, sheet).Range(ref)

            if options.get("bold") is not None:
                rng.Font.Bold = options["bold"]
            if options.get("italic") is not None:
                rng.Font.Italic = options["italic"]
            if options.get("underline") is not None:
                rng.Font.Underline = _XL_UNDERLINE_SINGLE if options["underline"] else _XL_UNDERLINE_NONE
            if options.get("font_name"):
                rng.Font.Name = options["font_name"]
            if options.get("font_size"):
                rng.Font.Size = options["font_size"]
            if options.get("font_color"):
                rng.Font.Color = hex_to_bgr(options["font_color"])
            if options.get("fill_color"):
                rng.Interior.Color = hex_to_bgr(options["fill_color"])
            if options.get("number_format"):
                rng.NumberFormat = options["number_format"]
            if options.get("horizontal_alignment"):
                rng.HorizontalAlignment = _H_ALIGN.get(
                    options["horizontal_alignment"].lower(), -4131
                )
            if options.get("vertical_alignment"):
                rng.VerticalAlignment = _V_ALIGN.get(
                    options["vertical_alignment"].lower(), -4108
                )
            if options.get("wrap_text") is not None:
                rng.WrapText = options["wrap_text"]
            if options.get("border"):
                for edge in _BORDER_EDGES:
                    try:
                        b = rng.Borders(edge)
                        b.LineStyle = 1
                        b.Weight = 2
                    except Exception:
                        pass

        run_on_main_thread(_format)
        return "OK", False
    except Exception:
        return "ERROR: " + traceback.format_exc(), True


register(
    "pyxll_format_cells",
    (
        "Apply formatting to a range. Pass an 'options' object with any combination of: "
        "bold (bool), italic (bool), underline (bool), font_name (str), font_size (number), "
        "font_color ('#RRGGBB'), fill_color ('#RRGGBB'), number_format (str, e.g. '0.00%'), "
        "horizontal_alignment ('left'|'center'|'right'), vertical_alignment ('top'|'center'|'bottom'), "
        "wrap_text (bool), border (bool — adds a thin border on all edges). "
        "Unknown keys are ignored."
    ),
    {
        "type": "object",
        "properties": {
            "ref": {
                "type": "string",
                "description": "A1 range reference, e.g. 'A1:D5'.",
            },
            "options": {
                "type": "object",
                "description": "Formatting options — see tool description for supported keys.",
            },
            "sheet": {
                "type": "string",
                "description": "Worksheet name. Omit to use the active sheet.",
            },
        },
        "required": ["ref", "options"],
    },
    _impl,
)
