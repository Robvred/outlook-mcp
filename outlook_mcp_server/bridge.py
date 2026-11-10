"""Outbound reverse bridge: reach the local MCP server through a relay on a remote host.

The MCP server listens on the loopback interface only. This module opens ONE outbound
WebSocket connection (``wss://``, port 443) to a relay running on the remote host, then
answers the HTTP requests the relay forwards over that connection by replaying them against
the local server. The remote MCP client talks to the relay; nothing ever connects inbound to
the workstation. See ``docs/remote-access.md`` and ``relay/``.

Wire protocol (JSON text frames, bodies base64-encoded):

    relay  -> bridge  {"type": "request",  "id": "...", "method": "POST", "path": "/mcp",
                       "query": "a=b", "headers": {...}, "body": "<base64>"}
    bridge -> relay   {"type": "response", "id": "...", "status": 200,
                       "headers": {...}, "body": "<base64>"}

Guardrails:

- The bridge only replays requests to the configured loopback target.
- ``GET`` requests are answered locally with 405: the MCP Streamable HTTP spec allows it,
  and it prevents a client from holding a never-ending SSE stream through the bridge.
- Bodies are capped. Hop-by-hop headers are dropped in both directions.
- The bridge authenticates to the relay with its own token, in addition to whatever headers
  the reverse proxy in front of the relay requires. The MCP bearer token is still checked by
  the MCP server itself; the bridge does not look at it.
"""

import argparse
import asyncio
import base64
import json
import logging
import os
import random
import sys
import threading
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence

from .remote import DEFAULT_HOST, DEFAULT_PATH, DEFAULT_PORT, MIN_TOKEN_LENGTH, RemoteConfigError, is_loopback

logger = logging.getLogger("outlook_mcp_server.bridge")

ENV_URL = "OUTLOOK_MCP_BRIDGE_URL"
ENV_TOKEN = "OUTLOOK_MCP_BRIDGE_TOKEN"
ENV_HEADERS = "OUTLOOK_MCP_BRIDGE_HEADERS"
ENV_TARGET = "OUTLOOK_MCP_BRIDGE_TARGET"

MAX_BODY_BYTES = 16 * 1024 * 1024
REQUEST_TIMEOUT_SECONDS = 120.0
RECONNECT_MIN_SECONDS = 1.0
RECONNECT_MAX_SECONDS = 30.0

HOP_BY_HOP = {
    "connection", "keep-alive", "proxy-authenticate", "proxy-authorization", "te", "trailers",
    "transfer-encoding", "upgrade", "host", "content-length", "accept-encoding", "content-encoding",
}


@dataclass
class BridgeSettings:
    url: str
    token: str = field(repr=False)
    target: str = f"http://{DEFAULT_HOST}:{DEFAULT_PORT}"
    headers: Dict[str, str] = field(default_factory=dict, repr=False)

    def validate(self) -> None:
        if not self.url.startswith(("wss://", "ws://")):
            raise RemoteConfigError(f"Bridge URL must start with wss:// (or ws:// for local tests), got '{self.url}'.")
        if self.url.startswith("ws://") and not _url_host_is_loopback(self.url):
            raise RemoteConfigError("Unencrypted ws:// is only allowed towards the loopback interface. Use wss://.")
        if not self.token or len(self.token) < MIN_TOKEN_LENGTH:
            raise RemoteConfigError(
                f"Bridge token missing or shorter than {MIN_TOKEN_LENGTH} characters. Set {ENV_TOKEN}."
            )
        if not self.target.startswith("http://") or not _url_host_is_loopback(self.target):
            raise RemoteConfigError(
                f"Bridge target must be an http:// URL on the loopback interface, got '{self.target}'. "
                "The bridge only replays requests to the local MCP server."
            )


def _url_host_is_loopback(url: str) -> bool:
    from urllib.parse import urlsplit

    host = urlsplit(url).hostname or ""
    return is_loopback(host)


def parse_header_lines(raw: Optional[str]) -> Dict[str, str]:
    """Parse 'Name: value' pairs separated by newlines or semicolons."""
    headers: Dict[str, str] = {}
    if not raw:
        return headers
    for part in raw.replace("\n", ";").split(";"):
        part = part.strip()
        if not part:
            continue
        if ":" not in part:
            raise RemoteConfigError(f"Header must look like 'Name: value', got '{part}'.")
        name, value = part.split(":", 1)
        name, value = name.strip(), value.strip()
        if not name or not value:
            raise RemoteConfigError(f"Header must look like 'Name: value', got '{part}'.")
        headers[name] = value
    return headers


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="outlook-mcp-bridge",
        description="Connect the local Outlook MCP server to a remote relay over an outbound WebSocket.",
    )
    parser.add_argument("--url", default=None, help=f"Relay WebSocket URL, e.g. wss://tunnel.example.com/bridge. Env: {ENV_URL}.")
    parser.add_argument("--token", default=None, help=f"Bridge token shared with the relay. Prefer env {ENV_TOKEN}.")
    parser.add_argument(
        "--header", action="append", default=[], metavar="'Name: value'",
        help=f"Extra header sent on the WebSocket handshake (reverse-proxy access tokens). Repeatable. Env: {ENV_HEADERS} ('Name: v; Other: w').",
    )
    parser.add_argument("--target", default=None, help=f"Local MCP server base URL. Default http://{DEFAULT_HOST}:{DEFAULT_PORT}. Env: {ENV_TARGET}.")
    return parser


def load_bridge_settings(argv: Optional[Sequence[str]] = None, environ: Optional[dict] = None) -> BridgeSettings:
    env = os.environ if environ is None else environ
    args = build_parser().parse_args(list(argv) if argv is not None else None)
    url = args.url or env.get(ENV_URL) or ""
    if not url:
        raise RemoteConfigError(f"Relay URL missing. Set {ENV_URL} or pass --url.")
    headers = parse_header_lines(env.get(ENV_HEADERS))
    for line in args.header:
        headers.update(parse_header_lines(line))
    settings = BridgeSettings(
        url=url,
        token=args.token or env.get(ENV_TOKEN) or "",
        target=args.target or env.get(ENV_TARGET) or f"http://{DEFAULT_HOST}:{DEFAULT_PORT}",
        headers=headers,
    )
    settings.validate()
    return settings


def bridge_settings_from_env(environ: Optional[dict] = None) -> Optional[BridgeSettings]:
    """Return settings if the bridge is configured through the environment, else None."""
    env = os.environ if environ is None else environ
    if not env.get(ENV_URL):
        return None
    return load_bridge_settings([], env)


def _clean_headers(headers: Dict[str, str]) -> Dict[str, str]:
    return {k: v for k, v in headers.items() if k.lower() not in HOP_BY_HOP}


def _error_response(req_id: str, status: int, message: str) -> dict:
    body = json.dumps({"error": message}).encode("utf-8")
    return {
        "type": "response", "id": req_id, "status": status,
        "headers": {"content-type": "application/json"},
        "body": base64.b64encode(body).decode("ascii"),
    }


async def handle_request(message: dict, client, target: str) -> dict:
    """Replay one forwarded request against the local server and build the response frame."""
    req_id = str(message.get("id", ""))
    method = str(message.get("method", "GET")).upper()
    path = str(message.get("path", "/"))
    query = str(message.get("query", "") or "")

    if method == "GET":
        return _error_response(req_id, 405, "GET is not relayed through the bridge")
    if not path.startswith("/"):
        return _error_response(req_id, 400, "invalid path")

    try:
        body = base64.b64decode(message.get("body", "") or "")
    except (ValueError, TypeError):
        return _error_response(req_id, 400, "invalid body encoding")
    if len(body) > MAX_BODY_BYTES:
        return _error_response(req_id, 413, "request body too large")

    headers = _clean_headers({str(k): str(v) for k, v in (message.get("headers") or {}).items()})
    headers["accept-encoding"] = "identity"
    url = target.rstrip("/") + path + (f"?{query}" if query else "")

    try:
        response = await client.request(method, url, headers=headers, content=body, timeout=REQUEST_TIMEOUT_SECONDS)
    except Exception as exc:  # local server down, timeout, ...
        logger.warning("Local request failed: %s", exc)
        return _error_response(req_id, 502, "local MCP server unreachable")

    content = response.content
    if len(content) > MAX_BODY_BYTES:
        return _error_response(req_id, 502, "response body too large")
    return {
        "type": "response", "id": req_id, "status": response.status_code,
        "headers": _clean_headers(dict(response.headers)),
        "body": base64.b64encode(content).decode("ascii"),
    }


async def serve_connection(ws, client, target: str) -> None:
    """Answer forwarded requests on one established WebSocket until it closes."""
    tasks = set()

    async def one(raw):
        try:
            message = json.loads(raw)
        except (ValueError, TypeError):
            logger.warning("Dropping non-JSON frame from relay")
            return
        if message.get("type") != "request":
            return
        reply = await handle_request(message, client, target)
        await ws.send(json.dumps(reply))

    try:
        async for raw in ws:
            task = asyncio.ensure_future(one(raw))
            tasks.add(task)
            task.add_done_callback(tasks.discard)
    finally:
        for task in tasks:
            task.cancel()


async def run_bridge(settings: BridgeSettings, stop: Optional[asyncio.Event] = None) -> None:
    """Keep one connection to the relay open, reconnecting with backoff."""
    import httpx
    from websockets.asyncio.client import connect

    handshake_headers = dict(settings.headers)
    handshake_headers["Authorization"] = f"Bearer {settings.token}"
    delay = RECONNECT_MIN_SECONDS

    async with httpx.AsyncClient() as client:
        while not (stop and stop.is_set()):
            try:
                logger.info("Connecting to relay %s", settings.url)
                async with connect(
                    settings.url, additional_headers=handshake_headers,
                    ping_interval=20, ping_timeout=20, max_size=MAX_BODY_BYTES + 4096,
                ) as ws:
                    logger.info("Connected to relay; replaying requests to %s", settings.target)
                    delay = RECONNECT_MIN_SECONDS
                    await serve_connection(ws, client, settings.target)
                    logger.info("Relay connection closed")
            except asyncio.CancelledError:
                raise
            except Exception as exc:
                logger.warning("Relay connection failed: %s", exc)
            if stop and stop.is_set():
                break
            sleep_for = delay + random.uniform(0, delay / 2)
            logger.info("Reconnecting in %.0fs", sleep_for)
            await asyncio.sleep(sleep_for)
            delay = min(delay * 2, RECONNECT_MAX_SECONDS)


def start_bridge_thread(settings: BridgeSettings) -> threading.Thread:
    """Run the bridge in a daemon thread with its own event loop (used by the server process)."""

    def runner():
        asyncio.run(run_bridge(settings))

    thread = threading.Thread(target=runner, name="outlook-mcp-bridge", daemon=True)
    thread.start()
    return thread


def main(argv: Optional[Sequence[str]] = None) -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s", stream=sys.stderr)
    try:
        settings = load_bridge_settings(argv)
    except RemoteConfigError as exc:
        print(f"Configuration error: {exc}", file=sys.stderr)
        sys.exit(2)
    try:
        asyncio.run(run_bridge(settings))
    except KeyboardInterrupt:
        print("Bridge stopped by user.", file=sys.stderr)


if __name__ == "__main__":
    main()
