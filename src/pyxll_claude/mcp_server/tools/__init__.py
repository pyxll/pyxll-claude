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


def call_tool(name: str, args: dict, ctx: dict) -> tuple[str, bool] | None:
    """Call a registered tool; returns (text, is_error) or None if unknown."""
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
    get_selection,
    read_range,
    read_formulas,
    write_range,
)
