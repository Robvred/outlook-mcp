# Remote access through an outbound bridge

The server normally runs over stdio next to a local MCP client. This page describes how to
reach it from an MCP client running on a remote host (for example an agent on a VPS) under
these constraints:

- the only flow leaving the Windows workstation is outbound TCP 443;
- no admin rights, no driver, no Windows service, no firewall prompt, no third-party binary
  on the workstation (tunnel tools such as chisel or frp are flagged by antivirus engines);
- the connection is initiated by the workstation, never by the remote host;
- every hop is authenticated.

Nothing in this setup changes what the tools do: drafts only, recoverable deletes, local
audit log. It only changes who can call them, so read the security notes before enabling it.

## Architecture

```
remote agent ──(loopback 8765)──> relay ──(one outbound wss on 443)──> bridge ──(loopback 8765)──> MCP server
    VPS                            VPS            ▲                   workstation                workstation
                                                  │ through the reverse proxy, which checks its
                                                  │ own access-token headers before the relay
                                                  └── initiated by the workstation
```

Everything on the workstation side is Python, already required for the server:

1. The **MCP server** listens on `127.0.0.1:8765` in HTTP mode and requires a bearer token
   (`OUTLOOK_MCP_TOKEN`).
2. The **bridge** (`outlook_mcp_server/bridge.py`) opens one outbound WebSocket to the relay
   and replays the requests it receives against the local server. It runs inside the server
   process when `OUTLOOK_MCP_BRIDGE_URL` is set, or standalone with `outlook-mcp-bridge`.
3. The **relay** (`relay/`) runs on the VPS in a container. It accepts the bridge on one port,
   placed behind the reverse proxy on 443, and serves the agent on another port bound to the
   VPS loopback.
4. The **remote agent** talks to `http://127.0.0.1:8765/mcp` on the VPS with the MCP token.

Why not WireGuard-based connectors (Newt, Tailscale): their data plane is UDP. Why not chisel,
frp, rathole: Go tunnel binaries are detected as hack tools by antivirus engines. Why not a
hosted tunnel: mailbox content would transit a third party.

Three independent secrets protect the chain. Leaking one does not open the mailbox:

| Hop | Secret | Checked by |
|---|---|---|
| Reach the relay | Reverse-proxy access token (headers) | Reverse proxy |
| Attach the bridge | Bridge token (`BRIDGE_TOKEN` = `OUTLOOK_MCP_BRIDGE_TOKEN`) | Relay |
| Call a tool | MCP bearer token (`OUTLOOK_MCP_TOKEN`) | this server |

The relay and the bridge never read the MCP token; they only carry it.

## 1. Relay (VPS)

```bash
git clone https://github.com/Robvred/outlook-mcp.git && cd outlook-mcp/relay
cp docker-compose.example.yml docker-compose.yml   # set BRIDGE_TOKEN, attach to the proxy network
docker compose up -d --build
```

| Variable | Default | Meaning |
|---|---|---|
| `BRIDGE_TOKEN` | none, mandatory | At least 32 characters. Same value as `OUTLOOK_MCP_BRIDGE_TOKEN` on the workstation |
| `BRIDGE_PORT` | `8080` | Bridge listener (`/bridge` WebSocket, `/healthz`). Behind the reverse proxy |
| `AGENT_PORT` | `8765` | Agent listener. Publish on `127.0.0.1` only |
| `REQUEST_TIMEOUT_SECONDS` | `130` | How long the relay waits for the bridge's answer |

Create a resource on the reverse proxy, for example `tunnel.example.com` → relay port 8080,
protected by access-token headers. The proxy must pass WebSocket upgrades (Traefik-based
proxies do). Checks:

```bash
curl -s -o /dev/null -w '%{http_code}\n' https://tunnel.example.com/healthz
curl -s -H "P-Access-Token-Id: <id>" -H "P-Access-Token: <token>" https://tunnel.example.com/healthz
```

Expected: a refusal without headers; `{"bridge_connected": false}` with them (HTTP 503 until
the bridge is up). The header names are the reverse proxy's; the ones above are Pangolin
defaults.

## 2. MCP server and bridge (workstation)

Settings come from environment variables, overridable by command-line flags.

| Variable | Default | Meaning |
|---|---|---|
| `OUTLOOK_MCP_TRANSPORT` | `stdio` | `http` enables the HTTP listener |
| `OUTLOOK_MCP_HOST` | `127.0.0.1` | Non-loopback refused unless `OUTLOOK_MCP_ALLOW_NON_LOOPBACK=1` |
| `OUTLOOK_MCP_PORT` | `8765` | |
| `OUTLOOK_MCP_PATH` | `/mcp` | |
| `OUTLOOK_MCP_TOKEN` | none | Mandatory in HTTP mode, at least 32 characters |
| `OUTLOOK_MCP_BRIDGE_URL` | none | `wss://tunnel.example.com/bridge`. Starts the bridge in the server process |
| `OUTLOOK_MCP_BRIDGE_TOKEN` | none | Same value as the relay's `BRIDGE_TOKEN` |
| `OUTLOOK_MCP_BRIDGE_HEADERS` | none | Reverse-proxy headers, `Name: value; Other: value` |
| `OUTLOOK_MCP_BRIDGE_TARGET` | `http://127.0.0.1:8765` | Local server. Must be a loopback `http://` URL |
| `OUTLOOK_MCP_BRIDGE_CA` | none | PEM file with extra trusted root CAs for the relay connection |

Generate tokens (one for the MCP server, one shared with the relay):

```powershell
python -c "import secrets; print(secrets.token_urlsafe(32))"
```

Start everything in one process (Outlook must be running in the same session):

```powershell
$env:OUTLOOK_MCP_TRANSPORT = "http"
$env:OUTLOOK_MCP_TOKEN = "<mcp token>"
$env:OUTLOOK_MCP_BRIDGE_URL = "wss://tunnel.example.com/bridge"
$env:OUTLOOK_MCP_BRIDGE_TOKEN = "<bridge token>"
$env:OUTLOOK_MCP_BRIDGE_HEADERS = "P-Access-Token-Id: <id>; P-Access-Token: <token>"
python -m outlook_mcp_server
```

Expected on stderr: `Listening on http://127.0.0.1:8765/mcp (bearer token required)`,
`Starting outbound bridge to wss://...`, then `Connected to relay`. The bridge reconnects by
itself with backoff after a network cut.

The bridge can also run as its own process, with the same variables or flags:

```powershell
outlook-mcp-bridge --url wss://tunnel.example.com/bridge --header "P-Access-Token-Id: <id>" --header "P-Access-Token: <token>"
```

Put the variables in a `.cmd` or `.ps1` in a user folder (never in the repository) and start it
from a scheduled task "at logon" of the session where Outlook runs, or from the Startup
folder. Do not run it as a Windows service: Outlook COM needs the interactive session.

Behaviour and guardrails:

- The HTTP listener binds to the loopback interface, which does not trigger the Windows
  firewall prompt. Non-loopback binds are refused unless explicitly allowed.
- In HTTP mode the server answers in stateless JSON mode: every request is self-contained,
  which is what the relay path expects. `GET` on the MCP path is answered 405 by the bridge,
  as the Streamable HTTP spec allows, so no long-lived stream can be held through it.
- The bridge only replays requests to the configured loopback target. Bodies are capped at
  16 MB. Hop-by-hop headers are dropped.
- Prefer environment variables over `--token`: command-line arguments show in the process
  list.
- If the workstation uses an explicit HTTP proxy, set `HTTPS_PROXY`; the WebSocket client
  honours it.
- TLS verification is never disabled. On Windows the bridge verifies the relay's certificate
  against the operating system store (via `truststore`), the same roots curl.exe and the
  browser use, including corporate roots deployed by group policy and TLS-inspection
  proxies. If `Relay connection failed: ... CERTIFICATE_VERIFY_FAILED` still appears, export
  the missing root CA from the Windows store as Base64 PEM and point `OUTLOOK_MCP_BRIDGE_CA`
  at it.

## 3. Remote agent (VPS)

Example for an agent that reads MCP servers from a YAML file:

```yaml
mcp_servers:
  outlook:
    url: http://127.0.0.1:8765/mcp
    headers:
      Authorization: Bearer <mcp token>
```

Checks from the VPS:

```bash
curl -s http://127.0.0.1:8765/healthz
curl -s -o /dev/null -w '%{http_code}\n' -X POST -H 'Content-Type: application/json' -d '{}' http://127.0.0.1:8765/mcp
```

Expected: `{"bridge_connected": true}`, then `401` (the MCP server refused the missing token,
which proves the whole chain is up). Then let the agent list the tools and call
`mailbox_stats_tool`, which is read-only.

## Testing the chain step by step

Validate one hypothesis at a time; each step rules out one layer.

1. **Egress.** From the workstation, `curl.exe -v https://tunnel.example.com/healthz` must
   reach the proxy (any HTTP code, no timeout).
2. **Relay alone.** On the VPS, `curl http://127.0.0.1:8765/healthz` returns
   `{"bridge_connected": false}`.
3. **Proxy in front of the relay.** The two curls of section 1: refusal without headers, the
   health JSON with them. If both return 404, the proxy is not routing the host.
4. **Bridge with a dummy target.** On the workstation, start any loopback listener on 8765
   (a PowerShell `HttpListener` on `http://localhost:8765/` works without admin rights), run
   `outlook-mcp-bridge` with the headers, then on the VPS `curl -X POST http://127.0.0.1:8765/`
   must return the dummy response. No firewall or UAC prompt must have appeared. Cut the
   workstation network for 30 seconds and check the bridge reconnects.
5. **Real server.** Stop the dummy listener, start the server with the variables of section 2
   and run the checks of section 3.

## Security notes

- The remote agent gets the same capabilities as a local client: full read access to the
  mailbox, draft creation, moves, categories, flags, saved attachments. Restrict the tool set
  on the agent side if it only needs to read.
- Keep the reverse proxy up to date; past versions had authentication bypasses affecting
  header-based access tokens.
- Rotate the three secrets independently. Changing `OUTLOOK_MCP_TOKEN` or the bridge token
  requires restarting the server (and the relay for the bridge token).
- The relay accepts one bridge at a time; a new connection replaces the previous one. Only the
  holder of the bridge token can do that.
- The audit log stays on the workstation (`~/.outlook-mcp/audit.log`). Remote calls are logged
  like local ones; there is no caller identity beyond the shared token.
- If a secret is pasted anywhere outside the two machines (chat, ticket, screenshot), rotate it.
