"""
Minimal MCP (Model Context Protocol) server over Streamable HTTP.

Runs in a background daemon thread inside Excel's Python process and exposes
tools to the claude agent via the tool registry in .tools.

Transport: MCP protocol 2025-03-26 Streamable HTTP (POST /mcp).
Each request is a self-contained JSON-RPC call; responses are returned
directly in the HTTP response body as application/json.
"""
import http.server
import json
import logging
import socket
import sys
import threading
import traceback
from http import HTTPStatus
from pathlib import Path
from typing import Any

import pyxll

from . import tools

_log = logging.getLogger(__name__)

_MCP_VERSION = "2025-03-26"

# Key used in sys.modules to persist the singleton across PyXLL reloads.
_REGISTRY_KEY = "__pyxll_mcp_registry__"


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

    def start(self) -> bool:
        """Bind the server and start the background thread.

        Returns True on success, False if the port is already in use.
        """
        try:
            handler = self._make_handler()
            self._server = _ThreadingHTTPServer(("127.0.0.1", self._port), handler)
            self._thread = threading.Thread(
                target=self._server.serve_forever,
                name="pyxll-mcp-server",
                daemon=True,
            )
            self._thread.start()
            _log.debug("PyXLL MCP server started on port %d", self._port)
            return True
        except OSError as exc:
            _log.warning("PyXLL MCP server failed to start on port %d: %s", self._port, exc)
            return False

    def stop(self) -> None:
        if self._server:
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
                if self.path.split("?")[0] == "/health":
                    self._send_json({"status": "ok"})
                else:
                    self.send_error(HTTPStatus.NOT_FOUND)

            def do_POST(self):  # noqa: N802
                if self.path.split("?")[0] == "/mcp":
                    self._handle_mcp()
                else:
                    self.send_error(HTTPStatus.NOT_FOUND)

            def do_DELETE(self):  # noqa: N802
                # Session termination — acknowledged but we have no state to clean up.
                if self.path.split("?")[0] == "/mcp":
                    self.send_response(HTTPStatus.OK)
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                else:
                    self.send_error(HTTPStatus.NOT_FOUND)

            def _handle_mcp(self):
                length = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(length)

                try:
                    msg = json.loads(body)
                except json.JSONDecodeError:
                    self.send_error(HTTPStatus.BAD_REQUEST, "Invalid JSON")
                    return

                method = msg.get("method", "")
                if method.startswith("notifications/") or "id" not in msg:
                    self.send_response(HTTPStatus.ACCEPTED)
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    return

                response = server_ref._dispatch(msg)
                if response is None:
                    self.send_response(HTTPStatus.ACCEPTED)
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    return

                self._send_json(response)

            def _send_json(self, data: dict) -> None:
                body = json.dumps(data).encode()
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(body)))
                self.send_header("Access-Control-Allow-Origin", "*")
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
            return None

        if method == "tools/list":
            return {"jsonrpc": "2.0", "id": msg_id, "result": {"tools": tools.get_tools_list()}}

        if method == "tools/call":
            return self._handle_tool_call(msg)

        return _jsonrpc_error(msg_id, -32601, f"Method not found: {method}")

    def _handle_tool_call(self, msg: dict) -> dict:
        msg_id = msg.get("id")
        params = msg.get("params", {})
        tool_name = params.get("name", "")
        args = params.get("arguments", {})
        ctx = {"workspace": self._workspace}

        try:
            result = tools.call_tool(tool_name, args, ctx)
            if result is None:
                return _jsonrpc_error(msg_id, -32602, f"Unknown tool: {tool_name}")
            text, is_error = result
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


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _jsonrpc_error(msg_id: Any, code: int, message: str) -> dict:
    return {
        "jsonrpc": "2.0",
        "id": msg_id,
        "error": {"code": code, "message": message},
    }


# ---------------------------------------------------------------------------
# Global singleton — one server per Excel process, survives PyXLL reloads
# ---------------------------------------------------------------------------

def _registry() -> dict:
    reg = sys.modules.get(_REGISTRY_KEY)
    if reg is None:
        reg = {"server": None, "port": None, "lock": threading.Lock(), "on_change": None}
        sys.modules[_REGISTRY_KEY] = reg
    return reg


def get_global_port() -> "int | None":
    """Return the port the global MCP server is listening on, or None."""
    return _registry().get("port")


def set_on_server_change(fn) -> None:
    """Register a callback invoked (on the main thread) when the server starts or stops."""
    _registry()["on_change"] = fn


def get_or_start_global(
    workspace: Path,
    port_start: int = 54717,
    port_end: int = 54816,
) -> "tuple[PyXLLMCPServer, int] | None":
    """Return the running global MCP server, starting one if needed.

    Returns (server, port) on success, None if no free port was found.
    Thread-safe; at most one server is ever started.
    """
    reg = _registry()
    on_change = None
    with reg["lock"]:
        if reg["server"] is not None:
            return reg["server"], reg["port"]
        port = find_free_port(port_start, port_end)
        if port is None:
            _log.warning("get_or_start_global: no free port in %d–%d", port_start, port_end)
            return None
        server = PyXLLMCPServer(workspace=workspace, port=port)
        if not server.start():
            return None
        reg["server"] = server
        reg["port"] = port
        on_change = reg["on_change"]
    if on_change is not None:
        pyxll.schedule_call(on_change)
    return server, port


def stop_global() -> None:
    """Stop the global MCP server if one is running."""
    reg = _registry()
    on_change = None
    with reg["lock"]:
        server = reg["server"]
        if server is not None:
            server.stop()
            reg["server"] = None
            reg["port"] = None
            on_change = reg["on_change"]
    if on_change is not None:
        pyxll.schedule_call(on_change)


def write_mcp_json(workspace: Path, port: int) -> None:
    """Write .mcp.json into the workspace pointing at the local MCP server."""
    mcp_json = workspace / ".mcp.json"
    mcp_json.write_text(
        json.dumps(
            {
                "mcpServers": {
                    "pyxll": {
                        "type": "streamable-http",
                        "url": f"http://localhost:{port}/mcp",
                    }
                }
            },
            indent=2,
        ),
        encoding="utf-8",
    )
