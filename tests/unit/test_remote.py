"""Tests for the optional HTTP transport settings (outlook_mcp_server.remote)."""

import pytest

from outlook_mcp_server.remote import (
    DEFAULT_HOST,
    DEFAULT_PATH,
    DEFAULT_PORT,
    MIN_TOKEN_LENGTH,
    RemoteConfigError,
    RemoteSettings,
    is_loopback,
    load_settings,
    token_matches,
)

GOOD_TOKEN = "a" * MIN_TOKEN_LENGTH


class TestDefaults:
    def test_stdio_by_default(self):
        settings = load_settings([], {})
        assert settings.transport == "stdio"
        assert not settings.is_http

    def test_http_defaults(self):
        settings = load_settings(["--transport", "http"], {"OUTLOOK_MCP_TOKEN": GOOD_TOKEN})
        assert settings.is_http
        assert settings.host == DEFAULT_HOST
        assert settings.port == DEFAULT_PORT
        assert settings.path == DEFAULT_PATH

    def test_token_not_in_repr(self):
        settings = RemoteSettings(transport="http", token=GOOD_TOKEN)
        assert GOOD_TOKEN not in repr(settings)


class TestPrecedence:
    def test_env_is_used(self):
        env = {
            "OUTLOOK_MCP_TRANSPORT": "http",
            "OUTLOOK_MCP_TOKEN": GOOD_TOKEN,
            "OUTLOOK_MCP_PORT": "9000",
            "OUTLOOK_MCP_PATH": "/outlook",
        }
        settings = load_settings([], env)
        assert settings.is_http
        assert settings.port == 9000
        assert settings.path == "/outlook"

    def test_args_override_env(self):
        env = {"OUTLOOK_MCP_TRANSPORT": "http", "OUTLOOK_MCP_TOKEN": GOOD_TOKEN, "OUTLOOK_MCP_PORT": "9000"}
        settings = load_settings(["--port", "9100"], env)
        assert settings.port == 9100

    def test_stdio_ignores_missing_token(self):
        settings = load_settings(["--transport", "stdio"], {"OUTLOOK_MCP_TOKEN": ""})
        assert not settings.is_http


class TestValidation:
    def test_http_requires_token(self):
        with pytest.raises(RemoteConfigError, match="bearer token"):
            load_settings(["--transport", "http"], {})

    def test_short_token_rejected(self):
        with pytest.raises(RemoteConfigError, match="too short"):
            load_settings(["--transport", "http", "--token", "short"], {})

    def test_non_loopback_rejected_by_default(self):
        with pytest.raises(RemoteConfigError, match="non-loopback"):
            load_settings(["--transport", "http", "--host", "0.0.0.0"], {"OUTLOOK_MCP_TOKEN": GOOD_TOKEN})

    def test_non_loopback_allowed_with_flag(self):
        settings = load_settings(
            ["--transport", "http", "--host", "0.0.0.0", "--allow-non-loopback"],
            {"OUTLOOK_MCP_TOKEN": GOOD_TOKEN},
        )
        assert settings.host == "0.0.0.0"

    def test_non_loopback_allowed_with_env(self):
        env = {"OUTLOOK_MCP_TOKEN": GOOD_TOKEN, "OUTLOOK_MCP_ALLOW_NON_LOOPBACK": "1"}
        settings = load_settings(["--transport", "http", "--host", "192.0.2.10"], env)
        assert settings.host == "192.0.2.10"

    def test_bad_port(self):
        with pytest.raises(RemoteConfigError, match="integer"):
            load_settings([], {"OUTLOOK_MCP_TRANSPORT": "http", "OUTLOOK_MCP_TOKEN": GOOD_TOKEN, "OUTLOOK_MCP_PORT": "abc"})

    def test_port_out_of_range(self):
        with pytest.raises(RemoteConfigError, match="between 1 and 65535"):
            load_settings(["--transport", "http", "--port", "70000"], {"OUTLOOK_MCP_TOKEN": GOOD_TOKEN})

    def test_path_must_start_with_slash(self):
        with pytest.raises(RemoteConfigError, match="start with '/'"):
            load_settings(["--transport", "http", "--path", "mcp"], {"OUTLOOK_MCP_TOKEN": GOOD_TOKEN})


class TestHelpers:
    @pytest.mark.parametrize("host", ["127.0.0.1", "localhost", "LOCALHOST", "::1", "127.0.0.2"])
    def test_loopback(self, host):
        assert is_loopback(host)

    @pytest.mark.parametrize("host", ["0.0.0.0", "192.0.2.10", "example.com", "::"])
    def test_not_loopback(self, host):
        assert not is_loopback(host)

    def test_token_matches(self):
        assert token_matches(GOOD_TOKEN, GOOD_TOKEN)
        assert not token_matches(GOOD_TOKEN, GOOD_TOKEN[:-1] + "b")
        assert not token_matches(GOOD_TOKEN, None)
        assert not token_matches("", "")
