"""
Minimal MCP (Model Context Protocol) server over HTTP/SSE.

Runs in a background daemon thread inside Excel's Python process and exposes
four tools to the claude agent:

  pyxll_reload_module   — reload a Python module in the workspace and call pyxll.rebind()
  pyxll_get_log         — tail the PyXLL log file
  pyxll_get_version     — return the installed PyXLL version string
  pyxll_get_python_version — return the Python version string
  pyxll_get_config_path — return the full path of the pyxll.cfg file
  pyxll_get_config      — return the parsed contents of pyxll.cfg via pyxll.get_config()
  pyxll_get_selection   — return the current selection address and sheet name
  pyxll_read_range      — read cell values via the Excel COM API
  pyxll_read_formulas   — read cell formulas (Formula2) via the Excel COM API
  pyxll_write_range     — write values to an Excel range via the COM API

Transport: MCP protocol 2024-11-05 over SSE (GET /sse + POST /message).
Each SSE connection gets a session ID; POST /message?sessionId=X routes
responses back to that session's event stream.
"""
import http.server
import importlib
import io
import json
import logging
import os
import queue
import socket
import sys
import threading
import traceback
import uuid
from http import HTTPStatus
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

import pyxll
from pyxll import get_config, rebind, xl_app

_log = logging.getLogger(__name__)

_MCP_VERSION = "2024-11-05"


# ---------------------------------------------------------------------------
# Public helper — port discovery
# ---------------------------------------------------------------------------

def find_free_port(start: int = 54717, end: int = 54816) -> int | None:
    """Return the first free port in [start, end], or None if all are taken."""
    for port in range(start, end + 1):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 0)
                s.bind(("127.0.0.1", port))
                return port
        except OSError:
            continue
    return None


# ---------------------------------------------------------------------------
# MCP server
# ---------------------------------------------------------------------------

class _ThreadingHTTPServer(http.server.ThreadingHTTPServer):
    """Threading HTTPServer that suppresses common connection errors during shutdown."""

    def handle_error(self, request, client_address):
        # Suppress ConnectionResetError and BrokenPipeError as they are expected
        # when the client process (Claude CLI) is terminated during task pane closure.
        exc_type, _, _ = sys.exc_info()
        if exc_type is not None and issubclass(exc_type, (ConnectionResetError, BrokenPipeError)):
            return
        super().handle_error(request, client_address)


class PyXLLMCPServer:
    """MCP server running on localhost in a background daemon thread."""

    def __init__(self, workspace: Path, port: int) -> None:
        self._workspace = workspace
        self._port = port
        self._server: http.server.HTTPServer | None = None
        self._thread: threading.Thread | None = None
        self._sessions: dict[str, queue.Queue] = {}
        self._sessions_lock = threading.Lock()

    def start(self) -> bool:
        """Bind the server and start the background thread.

        Returns True on success, False if the port is already in use.
        """
        try:
            handler = self._make_handler()
            self._server = _ThreadingHTTPServer(
                ("127.0.0.1", self._port), handler
            )
            self._thread = threading.Thread(
                target=self._server.serve_forever,
                name="pyxll-mcp-server",
                daemon=True,
            )
            self._thread.start()
            _log.debug("PyXLL MCP server started on port %d", self._port)
            return True
        except OSError as exc:
            _log.warning(
                "PyXLL MCP server failed to start on port %d: %s", self._port, exc
            )
            return False

    def stop(self) -> None:
        if self._server:
            with self._sessions_lock:
                for q in self._sessions.values():
                    q.put(None)
            self._server.shutdown()
            self._server = None
            _log.debug("PyXLL MCP server stopped")

    # ------------------------------------------------------------------
    # Request handler factory
    # ------------------------------------------------------------------

    def _make_handler(self):
        server_ref = self

        class _MCPHandler(http.server.BaseHTTPRequestHandler):
            def log_message(self, fmt, *args):  # noqa: N802
                _log.debug("MCP %s " + fmt, self.address_string(), *args)

            def do_GET(self):  # noqa: N802
                path = self.path.split("?")[0]
                if path == "/sse":
                    self._handle_sse()
                elif path == "/health":
                    self._send_json({"status": "ok"})
                else:
                    self.send_error(HTTPStatus.NOT_FOUND)

            def do_POST(self):  # noqa: N802
                if self.path.split("?")[0] == "/message":
                    self._handle_message()
                else:
                    self.send_error(HTTPStatus.NOT_FOUND)

            # -- SSE stream -----------------------------------------------

            def _handle_sse(self):
                session_id = str(uuid.uuid4())
                q: queue.Queue = queue.Queue()
                with server_ref._sessions_lock:
                    server_ref._sessions[session_id] = q

                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "text/event-stream")
                self.send_header("Cache-Control", "no-cache")
                self.send_header("Connection", "keep-alive")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()

                # Tell the client where to POST messages for this session.
                self._write_sse("endpoint", f"/message?sessionId={session_id}")

                try:
                    while True:
                        try:
                            data = q.get(timeout=15)
                        except queue.Empty:
                            # Heartbeat keeps the TCP connection alive.
                            self.wfile.write(b": heartbeat\n\n")
                            self.wfile.flush()
                            continue
                        if data is None:  # shutdown signal
                            break
                        self._write_sse("message", data)
                except (BrokenPipeError, ConnectionResetError):
                    pass
                finally:
                    with server_ref._sessions_lock:
                        server_ref._sessions.pop(session_id, None)

            def _write_sse(self, event: str, data: str) -> None:
                msg = f"event: {event}\ndata: {data}\n\n".encode()
                self.wfile.write(msg)
                self.wfile.flush()

            # -- Incoming JSON-RPC ----------------------------------------

            def _handle_message(self):
                qs = parse_qs(urlparse(self.path).query)
                session_id = qs.get("sessionId", [None])[0]

                length = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(length)

                try:
                    msg = json.loads(body)
                except json.JSONDecodeError:
                    self.send_error(HTTPStatus.BAD_REQUEST, "Invalid JSON")
                    return

                # Acknowledge immediately; response travels via SSE.
                self.send_response(HTTPStatus.ACCEPTED)
                self.end_headers()

                response = server_ref._dispatch(msg)
                if response is not None:
                    with server_ref._sessions_lock:
                        q = server_ref._sessions.get(session_id)
                    if q:
                        q.put(json.dumps(response))
                    else:
                        _log.warning("MCP: no SSE session for sessionId=%s", session_id)

            def _send_json(self, data: dict) -> None:
                body = json.dumps(data).encode()
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

        return _MCPHandler

    # ------------------------------------------------------------------
    # JSON-RPC dispatch
    # ------------------------------------------------------------------

    def _dispatch(self, msg: dict) -> dict | None:
        method = msg.get("method", "")
        msg_id = msg.get("id")

        if method == "initialize":
            return {
                "jsonrpc": "2.0",
                "id": msg_id,
                "result": {
                    "protocolVersion": _MCP_VERSION,
                    "capabilities": {"tools": {}},
                    "serverInfo": {"name": "pyxll", "version": "1.0"},
                },
            }

        if method.startswith("notifications/"):
            return None  # notifications get no response

        if method == "tools/list":
            return {"jsonrpc": "2.0", "id": msg_id, "result": {"tools": _TOOLS}}

        if method == "tools/call":
            return self._handle_tool_call(msg)

        return _jsonrpc_error(msg_id, -32601, f"Method not found: {method}")

    def _handle_tool_call(self, msg: dict) -> dict:
        msg_id = msg.get("id")
        params = msg.get("params", {})
        tool_name = params.get("name", "")
        args = params.get("arguments", {})

        try:
            if tool_name == "pyxll_reload_module":
                text = self._tool_reload_module(args.get("module"))
                is_error = text != "OK"
            elif tool_name == "pyxll_get_version":
                text = self._tool_get_version()
                is_error = text.startswith("ERROR:")
            elif tool_name == "pyxll_get_python_version":
                text = self._tool_get_python_version()
                is_error = text.startswith("ERROR:")
            elif tool_name == "pyxll_get_config_path":
                text = self._tool_get_config_path()
                is_error = text.startswith("ERROR:")
            elif tool_name == "pyxll_get_config":
                text = self._tool_get_config()
                is_error = text.startswith("ERROR:")
            elif tool_name == "pyxll_get_log":
                text = self._tool_get_log(int(args.get("lines", 50)))
                is_error = False
            elif tool_name == "pyxll_get_selection":
                text = self._tool_get_selection()
                is_error = text.startswith("ERROR:")
            elif tool_name == "pyxll_read_range":
                text = self._tool_read_range(args.get("ref"), args.get("sheet"))
                is_error = text.startswith("ERROR:")
            elif tool_name == "pyxll_read_formulas":
                text = self._tool_read_formulas(args.get("ref"), args.get("sheet"))
                is_error = text.startswith("ERROR:")
            elif tool_name == "pyxll_write_range":
                text = self._tool_write_range(
                    args.get("ref"), args.get("values", []), args.get("sheet")
                )
                is_error = text != "OK"
            else:
                return _jsonrpc_error(msg_id, -32602, f"Unknown tool: {tool_name}")
        except Exception:
            text = traceback.format_exc()
            is_error = True
            _log.error("MCP tool %s raised unexpectedly: %s", tool_name, text)

        return {
            "jsonrpc": "2.0",
            "id": msg_id,
            "result": {
                "content": [{"type": "text", "text": text}],
                "isError": is_error,
            },
        }

    # ------------------------------------------------------------------
    # Tool implementations
    # ------------------------------------------------------------------

    def _tool_reload_module(self, module: str | None) -> str:
        """Reload a workspace module and rebind @xl_func functions with Excel."""
        module_name = module or "pyxll_claude_functions"
        result: list[str | None] = [None]
        done = threading.Event()
        ws_str = str(self._workspace)

        def _do_reload():
            try:
                if ws_str not in sys.path:
                    sys.path.insert(0, ws_str)
                if module_name in sys.modules:
                    importlib.reload(sys.modules[module_name])
                else:
                    importlib.import_module(module_name)
                rebind()
                result[0] = "OK"
            except Exception:
                result[0] = traceback.format_exc()
            finally:
                done.set()

        pyxll.schedule_call(_do_reload)
        timed_out = not done.wait(timeout=30)
        if timed_out:
            _log.warning("pyxll_reload_module(%s): timed out waiting for main-thread callback", module_name)
            return "ERROR: reload timed out (Excel may be busy)"
        return result[0]

    def _tool_get_version(self) -> str:
        """Return the installed PyXLL version string."""
        try:
            return pyxll.__version__
        except Exception:
            return "ERROR: " + traceback.format_exc()

    def _tool_get_python_version(self) -> str:
        """Return the Python version string."""
        try:
            return sys.version
        except Exception:
            return "ERROR: " + traceback.format_exc()

    def _tool_get_config_path(self) -> str:
        """Return the full path of the pyxll.cfg file."""
        try:
            env = os.environ.get("PYXLL_CONFIG_FILE", "").strip()
            if env:
                return env
            xll = Path(pyxll.__file__)
            return str(xll.with_suffix(".cfg"))
        except Exception:
            return "ERROR: " + traceback.format_exc()

    def _tool_get_config(self) -> str:
        """Return the parsed pyxll config as an INI-formatted string."""
        try:
            buf = io.StringIO()
            get_config().write(buf)
            return buf.getvalue()
        except Exception:
            return "ERROR: " + traceback.format_exc()

    def _tool_get_log(self, lines: int = 50) -> str:
        """Return the last N lines of the PyXLL log file."""
        try:
            log_path = _get_log_path()
            if not log_path.exists():
                return f"Log file not found: {log_path}"
            with open(log_path, encoding="utf-8", errors="replace") as f:
                f.seek(0, 2)
                size = f.tell()
                # Read a chunk from the end; 200 bytes/line is a safe estimate.
                chunk = min(size, lines * 200)
                f.seek(max(0, size - chunk))
                content = f.read()
            all_lines = content.splitlines()
            return "\n".join(all_lines[-lines:])
        except Exception:
            return "ERROR: " + traceback.format_exc()

    def _tool_get_selection(self) -> str:
        """Return the current selection address and sheet name as JSON."""
        result: list[Any] = [None]
        done = threading.Event()

        def _do_get():
            try:
                xl = xl_app()
                sel = xl.Selection
                if sel is None:
                    result[0] = "ERROR: no selection"
                    return
                address = sel.Address.replace("$", "")
                sheet = xl.ActiveSheet.Name
                result[0] = json.dumps({"address": address, "sheet": sheet})
            except Exception:
                result[0] = "ERROR: " + traceback.format_exc()
            finally:
                done.set()

        pyxll.schedule_call(_do_get)
        done.wait(timeout=15)

        if result[0] is None:
            return "ERROR: timed out"
        return result[0]

    def _tool_read_range(self, ref: str, sheet: str | None) -> str:
        """Read cell values from an Excel range; return JSON 2-D list."""
        result: list[Any] = [None]
        done = threading.Event()

        def _do_read():
            try:
                xl = xl_app()
                ws = xl.Sheets(sheet) if sheet else xl.ActiveSheet
                v = ws.Range(ref).Value2
                # Value2 returns a scalar for a single cell, tuple-of-tuples otherwise.
                if not isinstance(v, tuple):
                    v = ((v,),)
                result[0] = json.dumps([list(row) for row in v])
            except Exception:
                result[0] = "ERROR: " + traceback.format_exc()
            finally:
                done.set()

        pyxll.schedule_call(_do_read)
        done.wait(timeout=15)

        if result[0] is None:
            return "ERROR: read timed out"
        return result[0]

    def _tool_read_formulas(self, ref: str, sheet: str | None) -> str:
        """Read cell formulas (Formula2) from an Excel range; return JSON 2-D list."""
        result: list[Any] = [None]
        done = threading.Event()

        def _do_read():
            try:
                xl = xl_app()
                ws = xl.Sheets(sheet) if sheet else xl.ActiveSheet
                v = ws.Range(ref).Formula2
                if not isinstance(v, tuple):
                    v = ((v,),)
                result[0] = json.dumps([list(row) for row in v])
            except Exception:
                result[0] = "ERROR: " + traceback.format_exc()
            finally:
                done.set()

        pyxll.schedule_call(_do_read)
        done.wait(timeout=15)

        if result[0] is None:
            return "ERROR: read timed out"
        return result[0]

    def _tool_write_range(self, ref: str, values: list, sheet: str | None) -> str:
        """Write a 2-D list of values to an Excel range."""
        result: list[str | None] = [None]
        done = threading.Event()

        def _do_write():
            try:
                xl = xl_app()
                ws = xl.Sheets(sheet) if sheet else xl.ActiveSheet
                top_left = ws.Range(ref)
                rows = len(values)
                cols = max((len(r) for r in values), default=0)
                rng = top_left.GetResize(rows, cols)
                # Pad each row to uniform width as a list of lists.
                # win32com requires a true 2-D list-of-lists to produce the
                # correct SAFEARRAY shape; tuple-of-tuples collapses to 1-D
                # for single-column ranges and writes to the wrong cells.
                padded = [
                    list(row) + [None] * (cols - len(row)) for row in values
                ]
                has_formula = any(
                    isinstance(v, str) and v.startswith("=")
                    for row in padded for v in row
                )
                if has_formula:
                    for r_idx, row in enumerate(padded):
                        for c_idx, v in enumerate(row):
                            cell = top_left.GetOffset(r_idx, c_idx)
                            if isinstance(v, str) and v.startswith("="):
                                cell.Formula2 = v
                            else:
                                cell.Value2 = v
                else:
                    rng.Value2 = padded
                result[0] = "OK"
            except Exception:
                result[0] = "ERROR: " + traceback.format_exc()
            finally:
                done.set()

        pyxll.schedule_call(_do_write)
        done.wait(timeout=15)

        if result[0] is None:
            return "ERROR: write timed out"
        return result[0]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _jsonrpc_error(msg_id: Any, code: int, message: str) -> dict:
    return {
        "jsonrpc": "2.0",
        "id": msg_id,
        "error": {"code": code, "message": message},
    }


def _get_log_path() -> Path:
    """Return the PyXLL log file path from config, or the default location."""
    try:
        cfg = get_config()

        # [LOG] path — directory for log files.
        # [LOG] file — log filename (may include strftime patterns).
        directory = os.path.expandvars(cfg.get("LOG", "path").strip())
        filename  = os.path.expandvars(cfg.get("LOG", "file").strip())
        return Path(directory) / filename
    except Exception:
        _log.warning("_get_log_path: could not read log path from pyxll config", exc_info=True)
        return Path(os.environ.get("APPDATA", "")) / "PyXLL" / "pyxll.log"


# ---------------------------------------------------------------------------
# Tool definitions (MCP schema)
# ---------------------------------------------------------------------------

_TOOLS = [
    {
        "name": "pyxll_reload_module",
        "description": (
            "Reload a single Python module from the workspace and rebind all @xl_func "
            "functions with Excel. This reloads only the specified module — it does NOT "
            "perform a full PyXLL add-in reload (which would close the task pane). "
            "Call this after every edit to pyxll_claude_functions.py. "
            "Returns 'OK' on success or a Python traceback on error."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "module": {
                    "type": "string",
                    "description": (
                        "Name of the module to reload. "
                        "Defaults to 'pyxll_claude_functions' if omitted."
                    ),
                },
            },
            "required": [],
        },
    },
    {
        "name": "pyxll_get_log",
        "description": (
            "Return the last N lines from the PyXLL log file. "
            "Use this whenever the user asks to see the PyXLL log, check for errors or "
            "warnings, or investigate recent PyXLL activity. Also useful after a reload "
            "to catch errors that appear in the log but not in the reload traceback."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "lines": {
                    "type": "integer",
                    "description": "Number of lines to return from the end of the log. Defaults to 50.",
                }
            },
            "required": [],
        },
    },
    {
        "name": "pyxll_get_config_path",
        "description": (
            "Return the full path (including filename) of the pyxll.cfg configuration file. "
            "Checks the PYXLL_CONFIG_FILE environment variable first; if unset, derives the "
            "path from pyxll.__file__ by replacing its extension with .cfg (e.g. pyxll.xll "
            "→ pyxll.cfg). Use this before reading or editing pyxll.cfg."
        ),
        "inputSchema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "pyxll_get_version",
        "description": (
            "Return the installed PyXLL version string (e.g. '5.4.1'). "
            "Call this before using any PyXLL feature that specifies a minimum version "
            "in the documentation, to confirm the feature is available."
        ),
        "inputSchema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "pyxll_get_python_version",
        "description": (
            "Return the Python version string (e.g. '3.11.4 (main, ...')). "
            "Call this when you need to confirm Python version compatibility "
            "before using language features or library APIs that require a specific version."
        ),
        "inputSchema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "pyxll_get_config",
        "description": (
            "Return the complete, parsed contents of pyxll.cfg as an INI-formatted string. "
            "Uses pyxll.get_config() so variable substitutions and any externally merged "
            "config files are already applied. Prefer this over reading the raw file when "
            "you need to inspect or reason about PyXLL configuration values."
        ),
        "inputSchema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "pyxll_get_selection",
        "description": (
            "Return the current Excel selection as a JSON object with 'address' and 'sheet' fields. "
            "Call this first whenever you need to know which cell(s) the user has selected — "
            "e.g. 'check the current cell', 'read the selected range', or any read/write operation "
            "relative to the selection. Pass the returned address to pyxll_read_range or "
            "pyxll_write_range. Works for single cells, ranges, and non-contiguous selections."
        ),
        "inputSchema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "pyxll_read_range",
        "description": (
            "Read cell values from an Excel worksheet. "
            "Returns a JSON-encoded 2-D list (list of rows). "
            "Single cells return [[value]]."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "ref": {
                    "type": "string",
                    "description": "A1 range reference, e.g. 'A1' or 'B2:D5'.",
                },
                "sheet": {
                    "type": "string",
                    "description": "Worksheet name, e.g. 'Sheet1'. Omit to use the active sheet.",
                },
            },
            "required": ["ref"],
        },
    },
    {
        "name": "pyxll_read_formulas",
        "description": (
            "Read cell formulas from an Excel worksheet using the Formula2 property. "
            "Returns a JSON-encoded 2-D list (list of rows). "
            "Formula cells return the formula string (e.g. '=SUM(A1:A10)'); "
            "non-formula cells return the raw value. "
            "Single cells return [[value]]."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "ref": {
                    "type": "string",
                    "description": "A1 range reference, e.g. 'A1' or 'B2:D5'.",
                },
                "sheet": {
                    "type": "string",
                    "description": "Worksheet name, e.g. 'Sheet1'. Omit to use the active sheet.",
                },
            },
            "required": ["ref"],
        },
    },
    {
        "name": "pyxll_write_range",
        "description": (
            "Write values to an Excel worksheet range. "
            "Strings starting with '=' are written as formulas."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "ref": {
                    "type": "string",
                    "description": "Top-left cell of the destination range, e.g. 'B3'.",
                },
                "values": {
                    "type": "array",
                    "items": {"type": "array"},
                    "description": "2-D array of values (list of rows).",
                },
                "sheet": {
                    "type": "string",
                    "description": "Worksheet name. Omit to use the active sheet.",
                },
            },
            "required": ["ref", "values"],
        },
    },
]
