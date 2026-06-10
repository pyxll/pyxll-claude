"""
pyxll-claude: A PyXLL extension that embeds a Claude AI chat terminal into
Microsoft Excel as a Custom Task Pane.

Entry points are declared in pyproject.toml under [project.entry-points."pyxll"],
so PyXLL picks this package up automatically once it is installed — no changes
to pyxll.cfg are required.
"""

import configparser
import importlib.resources
import logging
import sys
from pathlib import Path

import pyxll
from pyxll import get_config

_log = logging.getLogger(__name__)


def pyxll_modules():
    """PyXLL entry point: return the list of module names for PyXLL to load.

    If the workspace is configured and pyxll_claude_functions.py already exists
    there, the workspace is added to sys.path and the module is included so
    previously written Excel functions are registered automatically on startup.
    """
    modules = ["pyxll_claude.xl_functions"]

    try:
        value = get_config().get("CLAUDE", "workspace").strip()
        workspace = Path(value) if value else None
    except (configparser.NoSectionError, configparser.NoOptionError):
        workspace = None
    except Exception:
        _log.warning("pyxll_claude: could not read [CLAUDE] workspace from pyxll.cfg", exc_info=True)
        workspace = None

    if workspace is None:
        workspace = Path(pyxll.__file__).parent / "pyxll_claude_workspace"

    if (workspace / "pyxll_claude_functions.py").exists():
        ws_str = str(workspace)
        if ws_str not in sys.path:
            sys.path.insert(0, ws_str)
        modules.append("pyxll_claude_functions")

    return modules


def pyxll_ribbon():
    """PyXLL entry point: return the ribbon XML for the Claude tab.

    Returns a list of (filename, xml_string) tuples so PyXLL can merge the XML
    with other ribbon configurations.
    """
    try:
        ribbon_xml = (
            importlib.resources.files("pyxll_claude.resources")
            .joinpath("ribbon.xml")
            .read_text(encoding="utf-8")
        )
        return [(None, ribbon_xml)]
    except Exception:
        _log.error("Failed to load ribbon.xml for pyxll_claude", exc_info=True)
        return []
