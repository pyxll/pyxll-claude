"""
pyxll-claude: A PyXLL extension that embeds a Claude AI chat terminal into
Microsoft Excel as a Custom Task Pane.

Entry points are declared in pyproject.toml under [project.entry-points."pyxll"],
so PyXLL picks this package up automatically once it is installed — no changes
to pyxll.cfg are required.
"""

import importlib.resources
import logging

_log = logging.getLogger(__name__)


def pyxll_modules():
    """PyXLL entry point: return the list of module names for PyXLL to load."""
    return [
        "pyxll_claude.xl_functions",
    ]


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
