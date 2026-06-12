"""
Tool registry for the PyXLL MCP server.

Each tool module calls register() when imported, adding itself to _registry.
Importing this package triggers all registrations via the bottom imports.
"""
from __future__ import annotations

from typing import Callable

_registry: dict[str, dict] = {}


def register(
    name: str,
    description: str,
    input_schema: dict,
    fn: Callable[[dict, dict], tuple[str, bool]],
) -> None:
    _registry[name] = {"description": description, "inputSchema": input_schema, "fn": fn}


def get_tools_list() -> list[dict]:
    return [
        {
            "name": name,
            "description": entry["description"],
            "inputSchema": entry["inputSchema"],
        }
        for name, entry in _registry.items()
    ]


def call_tool(name: str, args: dict, ctx: dict) -> tuple[str | list[dict], bool] | None:
    """Call a registered tool; returns (content, is_error) or None if unknown.

    content is either a plain str (text) or a list of MCP content dicts
    (e.g. [{"type": "image", "data": "<b64>", "mimeType": "image/png"}]).
    """
    entry = _registry.get(name)
    if entry is None:
        return None
    return entry["fn"](args, ctx)


# Import all tool modules to trigger self-registration — order determines tools/list order.
from . import (  # noqa: E402
    reload_module,
    get_log,
    get_config_path,
    get_version,
    get_python_version,
    get_config,
    get_sheets,
    get_used_range,
    get_selection,
    get_cell_info,
    get_named_ranges,
    read_range,
    read_formulas,
    write_range,
    clear_range,
    insert_rows,
    insert_columns,
    merge_cells,
    add_sheet,
    name_range,
    save_workbook,
    format_cells,
    auto_fit_column,
    calculate,
    screenshot,
)
