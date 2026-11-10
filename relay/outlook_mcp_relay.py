"""Relay for the Outlook MCP Server bridge.

Runs on the remote host (VPS). Two listeners:

- ``BRIDGE_PORT`` (default 8080): accepts ONE WebSocket connection from the workstation
  bridge at ``/bridge``. Put it behind the reverse proxy on 443, which adds its own
  authentication. The bridge must also present ``Authorization: Bearer <BRIDGE_TOKEN>``.
- ``AGENT_PORT`` (default 8765): plain HTTP for the local MCP client (the agent). Every
  request is serialized, sent over the WebSocket, and the bridge's response is returned.
  Bind it to the loopback interface only (docker ``-p 127.0.0.1:8765:8765``).

The relay never looks at the MCP bearer token; it is forwarded to the MCP server, which
verifies it. When no bridge is connected the agent gets 503. When a new bridge connects, the
previous connection is closed: the workstation reconnected after a network cut.

Wire protocol: see ``outlook_mcp_server/bridge.py``.
"""

import asyncio
import base64
import hmac
import json
import logging
import os
import secrets
import sys
from typing import Dict, Optional

from aiohttp import WSMsgType, web

logger = logging.getLogger("outlook_mcp_relay")

MAX_BODY_BYTES = 16 * 1024 * 1024
REQUEST_TIMEOUT_SECONDS = float(os.getenv("REQUEST_TIMEOUT_SECONDS", "130"))
HOP_BY_HOP = {
    "connection", "keep-alive", "proxy-authenticate", "proxy-authorization", "te", "trailers",
    "transfer-encoding", "upgrade", "host", "content-length",
}


class RelayState:
    def __init__(self, bridge_token: str):
        self.bridge_token = bridge_token
        self.bridge: Optional[web.WebSocketResponse] = None
        self.pending: Dict[str, asyncio.Future] = {}
        self.lock = asyncio.Lock()

    def token_ok(self, header: Optional[str]) -> bool:
        if not header or not header.startswith("Bearer "):
            return False
        presented = header[len("Bearer "):].strip()
        return hmac.compare_digest(self.bridge_token.encode(), presented.encode())

    def fail_pending(self, reason: str) -> None:
        for fut in list(self.pending.values()):
            if not fut.done():
                fut.set_exception(ConnectionError(reason))
        self.pending.clear()


def _clean(headers) -> Dict[str, str]:
    return {k: v for k, v in headers.items() if k.lower() not in HOP_BY_HOP}


async def bridge_handler(request: web.Request) -> web.StreamResponse:
    state: RelayState = request.app["state"]
    if not state.token_ok(request.headers.get("Authorization")):
        logger.warning("Bridge handshake refused from %s", request.remote)
        return web.json_response({"error": "unauthorized"}, status=401)

    ws = web.WebSocketResponse(heartbeat=20, max_msg_size=MAX_BODY_BYTES + 4096)
    await ws.prepare(request)

    async with state.lock:
        previous = state.bridge
        state.bridge = ws
        state.fail_pending("bridge replaced")
    if previous is not None and not previous.closed:
        logger.info("Replacing previous bridge connection")
        await previous.close(code=4000, message=b"replaced")
    logger.info("Bridge connected from %s", request.remote)

    try:
        async for msg in ws:
            if msg.type != WSMsgType.TEXT:
                continue
            try:
                message = json.loads(msg.data)
            except ValueError:
                continue
            if message.get("type") != "response":
                continue
            fut = state.pending.pop(str(message.get("id", "")), None)
            if fut is not None and not fut.done():
                fut.set_result(message)
    finally:
        async with state.lock:
            if state.bridge is ws:
                state.bridge = None
                state.fail_pending("bridge disconnected")
        logger.info("Bridge disconnected")
    return ws


async def agent_handler(request: web.Request) -> web.StreamResponse:
    state: RelayState = request.app["state"]
    ws = state.bridge
    if ws is None or ws.closed:
        return web.json_response({"error": "bridge not connected"}, status=503)

    body = await request.read()
    if len(body) > MAX_BODY_BYTES:
        return web.json_response({"error": "request body too large"}, status=413)

    req_id = secrets.token_hex(8)
    frame = {
        "type": "request", "id": req_id, "method": request.method, "path": request.path,
        "query": request.query_string, "headers": _clean(request.headers),
        "body": base64.b64encode(body).decode("ascii"),
    }
    fut: asyncio.Future = asyncio.get_running_loop().create_future()
    state.pending[req_id] = fut
    try:
        await ws.send_str(json.dumps(frame))
        reply = await asyncio.wait_for(fut, timeout=REQUEST_TIMEOUT_SECONDS)
    except asyncio.TimeoutError:
        state.pending.pop(req_id, None)
        return web.json_response({"error": "bridge timeout"}, status=504)
    except ConnectionError as exc:
        return web.json_response({"error": str(exc)}, status=502)
    except Exception as exc:
        state.pending.pop(req_id, None)
        logger.warning("Forwarding failed: %s", exc)
        return web.json_response({"error": "bridge send failed"}, status=502)

    try:
        content = base64.b64decode(reply.get("body", "") or "")
    except (ValueError, TypeError):
        return web.json_response({"error": "invalid bridge response"}, status=502)
    headers = _clean({str(k): str(v) for k, v in (reply.get("headers") or {}).items()})
    return web.Response(status=int(reply.get("status", 502)), headers=headers, body=content)


async def health_handler(request: web.Request) -> web.Response:
    state: RelayState = request.app["state"]
    connected = state.bridge is not None and not state.bridge.closed
    return web.json_response({"bridge_connected": connected}, status=200 if connected else 503)


def build_apps(bridge_token: str):
    state = RelayState(bridge_token)
    bridge_app = web.Application(client_max_size=MAX_BODY_BYTES)
    bridge_app["state"] = state
    bridge_app.router.add_get("/bridge", bridge_handler)
    bridge_app.router.add_get("/healthz", health_handler)

    agent_app = web.Application(client_max_size=MAX_BODY_BYTES)
    agent_app["state"] = state
    agent_app.router.add_get("/healthz", health_handler)
    agent_app.router.add_route("*", "/{tail:.*}", agent_handler)
    return bridge_app, agent_app


async def serve(bridge_token: str, bridge_host: str, bridge_port: int, agent_host: str, agent_port: int) -> None:
    bridge_app, agent_app = build_apps(bridge_token)
    runners = []
    for app, host, port, name in ((bridge_app, bridge_host, bridge_port, "bridge"), (agent_app, agent_host, agent_port, "agent")):
        runner = web.AppRunner(app, access_log=None)
        await runner.setup()
        await web.TCPSite(runner, host, port).start()
        runners.append(runner)
        logger.info("%s listener on %s:%d", name, host, port)
    try:
        await asyncio.Event().wait()
    finally:
        for runner in runners:
            await runner.cleanup()


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s", stream=sys.stderr)
    token = os.getenv("BRIDGE_TOKEN", "")
    if len(token) < 32:
        print("BRIDGE_TOKEN missing or shorter than 32 characters.", file=sys.stderr)
        sys.exit(2)
    try:
        asyncio.run(serve(
            token,
            os.getenv("BRIDGE_HOST", "0.0.0.0"), int(os.getenv("BRIDGE_PORT", "8080")),
            os.getenv("AGENT_HOST", "0.0.0.0"), int(os.getenv("AGENT_PORT", "8765")),
        ))
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
