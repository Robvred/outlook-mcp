"""Tests for the outbound bridge (outlook_mcp_server.bridge)."""

import asyncio
import base64
import json

import pytest

from outlook_mcp_server.bridge import (
    BridgeSettings,
    bridge_settings_from_env,
    handle_request,
    load_bridge_settings,
    parse_header_lines,
)
from outlook_mcp_server.remote import RemoteConfigError

TOKEN = "b" * 32
URL = "wss://tunnel.example.com/bridge"


class TestHeaderParsing:
    def test_empty(self):
        assert parse_header_lines(None) == {}
        assert parse_header_lines("") == {}

    def test_semicolon_and_newline_separated(self):
        raw = "P-Access-Token-Id: abc; P-Access-Token: def\nX-Other: 1"
        assert parse_header_lines(raw) == {"P-Access-Token-Id": "abc", "P-Access-Token": "def", "X-Other": "1"}

    def test_value_may_contain_colon(self):
        assert parse_header_lines("Authorization: Bearer a:b") == {"Authorization": "Bearer a:b"}

    @pytest.mark.parametrize("raw", ["novalue", "Name:", ": value"])
    def test_invalid(self, raw):
        with pytest.raises(RemoteConfigError):
            parse_header_lines(raw)


class TestSettings:
    def test_from_args(self):
        s = load_bridge_settings(["--url", URL, "--token", TOKEN, "--header", "A: 1", "--header", "B: 2"], {})
        assert s.url == URL
        assert s.headers == {"A": "1", "B": "2"}
        assert s.target == "http://127.0.0.1:8765"

    def test_from_env_with_arg_override(self):
        env = {
            "OUTLOOK_MCP_BRIDGE_URL": URL,
            "OUTLOOK_MCP_BRIDGE_TOKEN": TOKEN,
            "OUTLOOK_MCP_BRIDGE_HEADERS": "A: env; B: env",
            "OUTLOOK_MCP_BRIDGE_TARGET": "http://localhost:9000",
        }
        s = load_bridge_settings(["--header", "B: arg"], env)
        assert s.headers == {"A": "env", "B": "arg"}
        assert s.target == "http://localhost:9000"

    def test_env_helper_returns_none_when_unset(self):
        assert bridge_settings_from_env({}) is None

    def test_env_helper_returns_settings(self):
        s = bridge_settings_from_env({"OUTLOOK_MCP_BRIDGE_URL": URL, "OUTLOOK_MCP_BRIDGE_TOKEN": TOKEN})
        assert s is not None and s.url == URL

    def test_url_required(self):
        with pytest.raises(RemoteConfigError, match="URL missing"):
            load_bridge_settings([], {"OUTLOOK_MCP_BRIDGE_TOKEN": TOKEN})

    def test_token_required_and_long(self):
        with pytest.raises(RemoteConfigError, match="token"):
            load_bridge_settings(["--url", URL], {})
        with pytest.raises(RemoteConfigError, match="token"):
            load_bridge_settings(["--url", URL, "--token", "short"], {})

    def test_plain_ws_only_to_loopback(self):
        BridgeSettings(url="ws://127.0.0.1:8080/bridge", token=TOKEN).validate()
        with pytest.raises(RemoteConfigError, match="wss"):
            BridgeSettings(url="ws://tunnel.example.com/bridge", token=TOKEN).validate()

    def test_target_must_be_loopback_http(self):
        with pytest.raises(RemoteConfigError, match="loopback"):
            BridgeSettings(url=URL, token=TOKEN, target="http://192.0.2.10:8765").validate()
        with pytest.raises(RemoteConfigError, match="loopback"):
            BridgeSettings(url=URL, token=TOKEN, target="https://127.0.0.1:8765").validate()

    def test_secrets_not_in_repr(self):
        s = BridgeSettings(url=URL, token=TOKEN, headers={"P-Access-Token": "secret-value"})
        assert TOKEN not in repr(s)
        assert "secret-value" not in repr(s)


class FakeResponse:
    def __init__(self, status_code=200, headers=None, content=b""):
        self.status_code = status_code
        self.headers = headers or {}
        self.content = content


class FakeClient:
    def __init__(self, response=None, error=None):
        self.response = response or FakeResponse()
        self.error = error
        self.calls = []

    async def request(self, method, url, headers=None, content=b"", timeout=None):
        self.calls.append({"method": method, "url": url, "headers": headers, "content": content})
        if self.error:
            raise self.error
        return self.response


def frame(method="POST", path="/mcp", body=b"{}", headers=None, query=""):
    return {
        "type": "request", "id": "42", "method": method, "path": path, "query": query,
        "headers": headers or {}, "body": base64.b64encode(body).decode("ascii"),
    }


def decode(reply):
    return json.loads(base64.b64decode(reply["body"]).decode("utf-8")) if reply["body"] else None


class TestHandleRequest:
    def test_replays_post_to_target(self):
        client = FakeClient(FakeResponse(200, {"content-type": "application/json", "transfer-encoding": "chunked"}, b'{"ok":1}'))
        reply = asyncio.run(handle_request(
            frame(body=b'{"jsonrpc":"2.0"}', query="x=1",
                  headers={"Authorization": "Bearer t", "Host": "relay", "Content-Length": "17", "Accept-Encoding": "gzip"}),
            client, "http://127.0.0.1:8765",
        ))
        call = client.calls[0]
        assert call["method"] == "POST"
        assert call["url"] == "http://127.0.0.1:8765/mcp?x=1"
        assert call["content"] == b'{"jsonrpc":"2.0"}'
        assert call["headers"]["Authorization"] == "Bearer t"
        assert "Host" not in call["headers"] and "Content-Length" not in call["headers"]
        assert call["headers"]["accept-encoding"] == "identity"
        assert reply["type"] == "response" and reply["id"] == "42" and reply["status"] == 200
        assert reply["headers"] == {"content-type": "application/json"}
        assert decode(reply) == {"ok": 1}

    def test_get_is_refused_locally(self):
        client = FakeClient()
        reply = asyncio.run(handle_request(frame(method="GET"), client, "http://127.0.0.1:8765"))
        assert reply["status"] == 405
        assert client.calls == []

    def test_invalid_path(self):
        reply = asyncio.run(handle_request(frame(path="mcp"), FakeClient(), "http://127.0.0.1:8765"))
        assert reply["status"] == 400

    def test_bad_body_encoding(self):
        bad = frame()
        bad["body"] = "%%%not base64%%%"
        reply = asyncio.run(handle_request(bad, FakeClient(), "http://127.0.0.1:8765"))
        assert reply["status"] == 400

    def test_local_server_down(self):
        reply = asyncio.run(handle_request(frame(), FakeClient(error=ConnectionError("refused")), "http://127.0.0.1:8765"))
        assert reply["status"] == 502
        assert decode(reply) == {"error": "local MCP server unreachable"}

    def test_status_passthrough(self):
        reply = asyncio.run(handle_request(frame(), FakeClient(FakeResponse(401, {}, b"")), "http://127.0.0.1:8765"))
        assert reply["status"] == 401
