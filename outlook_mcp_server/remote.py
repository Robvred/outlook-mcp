"""Remote access settings for the Outlook MCP Server.

By default the server speaks MCP over stdio to a client running on the same
machine. This module adds an optional HTTP transport, protected by a bearer
token, so that the server can be reached through an outbound reverse tunnel
(see ``docs/remote-access.md``). Nothing here changes what the tools do.

Design constraints:

- The HTTP listener binds to the loopback interface unless explicitly told
  otherwise. Binding to a non-loopback address on Windows triggers a firewall
  prompt and exposes the mailbox on the LAN, which is never wanted here.
- A token is mandatory in HTTP mode and must be long enough to be
  unguessable. Tokens are compared in constant time.
- Settings come from environment variables, overridable by command-line
  arguments, so that a scheduled task or a shortcut can start the server
  without a wrapper script.
"""

import argparse
import hmac
import ipaddress
import os
from dataclasses import dataclass, field
from typing import Optional, Sequence

ENV_TRANSPORT = "OUTLOOK_MCP_TRANSPORT"
ENV_HOST = "OUTLOOK_MCP_HOST"
ENV_PORT = "OUTLOOK_MCP_PORT"
ENV_PATH = "OUTLOOK_MCP_PATH"
ENV_TOKEN = "OUTLOOK_MCP_TOKEN"
ENV_ALLOW_NON_LOOPBACK = "OUTLOOK_MCP_ALLOW_NON_LOOPBACK"

DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8765
DEFAULT_PATH = "/mcp"
MIN_TOKEN_LENGTH = 32

TRANSPORTS = ("stdio", "http")


class RemoteConfigError(ValueError):
    """Raised when the remote access settings are unusable."""


@dataclass
class RemoteSettings:
    """Resolved transport settings."""

    transport: str = "stdio"
    host: str = DEFAULT_HOST
    port: int = DEFAULT_PORT
    path: str = DEFAULT_PATH
    token: Optional[str] = field(default=None, repr=False)
    allow_non_loopback: bool = False

    @property
    def is_http(self) -> bool:
        return self.transport == "http"

    def validate(self) -> None:
        """Check the settings; raise RemoteConfigError on the first problem."""
        if self.transport not in TRANSPORTS:
            raise RemoteConfigError(
                f"Unknown transport '{self.transport}'. Expected one of: {', '.join(TRANSPORTS)}."
            )
        if not self.is_http:
            return

        if not (1 <= self.port <= 65535):
            raise RemoteConfigError(f"Port must be between 1 and 65535, got {self.port}.")

        if not self.path.startswith("/"):
            raise RemoteConfigError(f"Path must start with '/', got '{self.path}'.")

        if not self.token:
            raise RemoteConfigError(
                f"HTTP transport requires a bearer token. Set {ENV_TOKEN} or pass --token."
            )
        if len(self.token) < MIN_TOKEN_LENGTH:
            raise RemoteConfigError(
                f"Bearer token is too short ({len(self.token)} characters). "
                f"Use at least {MIN_TOKEN_LENGTH}, for example the output of "
                "'python -c \"import secrets; print(secrets.token_urlsafe(32))\"'."
            )

        if not is_loopback(self.host) and not self.allow_non_loopback:
            raise RemoteConfigError(
                f"Refusing to bind to non-loopback address '{self.host}'. The server is meant to "
                "be reached through an outbound tunnel from 127.0.0.1. Set "
                f"{ENV_ALLOW_NON_LOOPBACK}=1 or pass --allow-non-loopback if you really want this."
            )


def is_loopback(host: str) -> bool:
    """Return True if host designates the local machine only."""
    if host.lower() == "localhost":
        return True
    try:
        return ipaddress.ip_address(host).is_loopback
    except ValueError:
        return False


def token_matches(expected: str, presented: Optional[str]) -> bool:
    """Constant-time comparison of a presented token against the configured one."""
    if not expected or not presented:
        return False
    return hmac.compare_digest(expected.encode("utf-8"), presented.encode("utf-8"))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="outlook-mcp-server",
        description=(
            "Outlook MCP Server. Default transport is stdio. Use --transport http to expose "
            "a token-protected Streamable HTTP endpoint on the loopback interface."
        ),
    )
    parser.add_argument(
        "--transport",
        choices=TRANSPORTS,
        default=None,
        help=f"stdio (default) or http. Env: {ENV_TRANSPORT}.",
    )
    parser.add_argument("--host", default=None, help=f"Bind address, default {DEFAULT_HOST}. Env: {ENV_HOST}.")
    parser.add_argument("--port", type=int, default=None, help=f"Bind port, default {DEFAULT_PORT}. Env: {ENV_PORT}.")
    parser.add_argument("--path", default=None, help=f"HTTP path, default {DEFAULT_PATH}. Env: {ENV_PATH}.")
    parser.add_argument(
        "--token",
        default=None,
        help=f"Bearer token required on every HTTP request. Prefer the env var {ENV_TOKEN} "
        "so the token does not show up in the process list.",
    )
    parser.add_argument(
        "--allow-non-loopback",
        action="store_true",
        default=False,
        help=f"Allow binding to a non-loopback address. Env: {ENV_ALLOW_NON_LOOPBACK}=1.",
    )
    return parser


def load_settings(argv: Optional[Sequence[str]] = None, environ: Optional[dict] = None) -> RemoteSettings:
    """Resolve settings from environment variables, then command-line overrides, and validate."""
    env = os.environ if environ is None else environ
    args = build_parser().parse_args(list(argv) if argv is not None else None)

    def pick(arg_value, env_name, default):
        if arg_value is not None:
            return arg_value
        env_value = env.get(env_name)
        if env_value is not None and env_value != "":
            return env_value
        return default

    port_raw = pick(args.port, ENV_PORT, DEFAULT_PORT)
    try:
        port = int(port_raw)
    except (TypeError, ValueError):
        raise RemoteConfigError(f"Port must be an integer, got '{port_raw}'.")

    allow_non_loopback = args.allow_non_loopback or (
        env.get(ENV_ALLOW_NON_LOOPBACK, "").strip().lower() in ("1", "true", "yes", "on")
    )

    settings = RemoteSettings(
        transport=str(pick(args.transport, ENV_TRANSPORT, "stdio")).lower(),
        host=str(pick(args.host, ENV_HOST, DEFAULT_HOST)),
        port=port,
        path=str(pick(args.path, ENV_PATH, DEFAULT_PATH)),
        token=pick(args.token, ENV_TOKEN, None),
        allow_non_loopback=allow_non_loopback,
    )
    settings.validate()
    return settings


def build_token_verifier(token: str):
    """Build a FastMCP token verifier that accepts exactly one static bearer token.

    Imported lazily so that the settings helpers above stay importable without fastmcp.
    """
    from fastmcp.server.auth import AccessToken, TokenVerifier

    class SingleTokenVerifier(TokenVerifier):
        """Accept one pre-shared token, compared in constant time."""

        def __init__(self, expected: str):
            super().__init__()
            self._expected = expected

        async def verify_token(self, token: str):  # type: ignore[override]
            if token_matches(self._expected, token):
                return AccessToken(token=token, client_id="outlook-mcp-remote", scopes=[], expires_at=None)
            return None

    return SingleTokenVerifier(token)
