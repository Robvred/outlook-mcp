# Remote access through an outbound tunnel

The server normally runs over stdio next to a local MCP client. This page describes how to
reach it from an MCP client running on a remote host (for example an agent on a VPS) under
these constraints:

- the only flow leaving the Windows workstation is outbound TCP 443;
- no admin rights, no driver, no Windows service, no firewall prompt on the workstation;
- the connection is initiated by the workstation, never by the remote host;
- every hop is authenticated.

Nothing in this setup changes what the tools do: drafts only, recoverable deletes, local
audit log. It only changes who can call them, so read the security notes before enabling it.

## Architecture

```
remote agent ──(loopback)──> chisel server ──(reverse tunnel)──> chisel client ──(loopback)──> MCP server
   VPS                           VPS                               workstation                workstation
                                  ▲
                                  │ wss:// on 443, through the reverse proxy
                                  │ (authenticated by the proxy, e.g. Pangolin access-token headers)
                                  └──────────────────────────── outbound from the workstation
```

1. The MCP server listens on `127.0.0.1:8765` on the workstation, in HTTP mode, and requires
   a bearer token.
2. A [chisel](https://github.com/jpillora/chisel) client on the workstation opens one
   outbound WebSocket-over-HTTPS connection to the VPS and publishes the local port 8765 on
   the VPS loopback (`R:127.0.0.1:8765:localhost:8765`).
3. The chisel server on the VPS sits behind the existing reverse proxy on 443. The proxy
   authenticates the tunnel connection (Pangolin resource with access-token headers, or any
   equivalent). The chisel server is never exposed directly.
4. The remote agent talks to `http://127.0.0.1:8765/mcp` on the VPS, with the MCP bearer token.

Why not WireGuard-based connectors (Newt, Tailscale, plain WireGuard): their data plane is
UDP, which does not fit a 443-only egress policy. Why not a hosted tunnel: mailbox content
would transit a third party.

Three independent secrets protect the chain. Leaking one does not open the mailbox:

| Hop | Secret | Checked by |
|---|---|---|
| Reach the chisel server | Reverse-proxy access token (headers) | Reverse proxy |
| Open the tunnel | chisel `--auth user:password` + server fingerprint | chisel |
| Call a tool | MCP bearer token (`OUTLOOK_MCP_TOKEN`) | this server |

## 1. MCP server in HTTP mode (workstation)

Settings come from environment variables, overridable by command-line flags:

| Variable | Flag | Default | Meaning |
|---|---|---|---|
| `OUTLOOK_MCP_TRANSPORT` | `--transport` | `stdio` | `stdio` or `http` |
| `OUTLOOK_MCP_HOST` | `--host` | `127.0.0.1` | Bind address. Non-loopback is refused unless `OUTLOOK_MCP_ALLOW_NON_LOOPBACK=1` |
| `OUTLOOK_MCP_PORT` | `--port` | `8765` | Bind port |
| `OUTLOOK_MCP_PATH` | `--path` | `/mcp` | HTTP path of the MCP endpoint |
| `OUTLOOK_MCP_TOKEN` | `--token` | none | Bearer token, mandatory in HTTP mode, at least 32 characters |

Generate a token:

```powershell
python -c "import secrets; print(secrets.token_urlsafe(32))"
```

Start the server (Outlook must be running in the same session):

```powershell
$env:OUTLOOK_MCP_TRANSPORT = "http"
$env:OUTLOOK_MCP_TOKEN = "<generated token>"
python -m outlook_mcp_server
```

Expected on stderr: `Listening on http://127.0.0.1:8765/mcp (bearer token required)`.

Binding to the loopback interface does not trigger the Windows firewall prompt. The server
refuses to start without a token or with a short one, and refuses non-loopback addresses
unless explicitly allowed. Prefer the environment variable over `--token`: command-line
arguments are visible in the process list.

Quick local check (expect `401` without the header and a JSON-RPC response with it):

```powershell
curl.exe -s -o NUL -w "%{http_code}`n" http://127.0.0.1:8765/mcp
```

## 2. chisel server (VPS)

Run it bound to the loopback or to the Docker network of the reverse proxy, never on a public
interface:

```bash
docker run -d --name chisel --restart unless-stopped -p 127.0.0.1:8080:8080 \
  jpillora/chisel server --port 8080 --reverse --auth tunnel:<chisel password>
```

`docker logs chisel` prints the server fingerprint. Keep it for the client.

Create a resource on the reverse proxy, for example `tunnel.example.com` → `chisel:8080`,
protected by access-token headers. Check from anywhere (expect a refusal without headers and a
`404` from chisel with them):

```bash
curl -s -o /dev/null -w '%{http_code}\n' https://tunnel.example.com
curl -s -o /dev/null -w '%{http_code}\n' -H "P-Access-Token-Id: <id>" -H "P-Access-Token: <token>" https://tunnel.example.com
```

The header names depend on the proxy; the ones above are Pangolin defaults. The proxy must pass
WebSocket upgrades to the target, which Traefik-based proxies do by default.

## 3. chisel client (workstation)

Download the Windows release binary into a user folder. No installation, no admin rights.

```powershell
chisel.exe client --auth tunnel:<chisel password> --fingerprint <server fingerprint> `
  --header "P-Access-Token-Id: <id>" --header "P-Access-Token: <token>" `
  --keepalive 25s https://tunnel.example.com R:127.0.0.1:8765:localhost:8765
```

The client reconnects by itself after a network cut. Start it, together with the MCP server,
from a scheduled task "at logon" of the user session where Outlook runs, or from a shortcut in
the Startup folder. Do not run either as a Windows service: Outlook COM needs the interactive
session.

If the workstation uses an explicit HTTP proxy, set `HTTPS_PROXY` before starting chisel.

## 4. Remote agent (VPS)

Example for an agent that reads MCP servers from a YAML file:

```yaml
mcp_servers:
  outlook:
    url: http://127.0.0.1:8765/mcp
    headers:
      Authorization: Bearer <OUTLOOK_MCP_TOKEN>
```

Check from the VPS before configuring the agent (expect `401`, which proves the whole chain
down to the MCP server is up):

```bash
curl -s -o /dev/null -w '%{http_code}\n' http://127.0.0.1:8765/mcp
```

Then let the agent list the tools and call `mailbox_stats_tool`, which is read-only.

## Testing the chain step by step

Validate one hypothesis at a time; each step rules out one layer.

1. **Egress and antivirus.** From the workstation, `curl.exe -v https://tunnel.example.com`
   must reach the proxy (any HTTP code, no timeout). Run `chisel.exe --version`; if the
   antivirus quarantines it, stop here.
2. **chisel server alone.** On the VPS, `curl http://127.0.0.1:8080` returns `404`.
3. **Proxy in front of chisel.** The two curls of section 2: refusal without headers, `404`
   with them. If both return `404`, the proxy is not routing the host (its own 404 page).
4. **Tunnel with a dummy target.** On the workstation, start any loopback listener on 8765
   (a PowerShell `HttpListener` on `http://localhost:8765/` works without admin rights), start
   the chisel client, then `curl http://127.0.0.1:8765/` on the VPS must return the dummy
   response. No firewall or UAC prompt must have appeared. Cut the workstation network for
   30 seconds and check the client reconnects.
5. **Real server.** Replace the dummy listener by the MCP server in HTTP mode and run the
   checks of section 4.

## Security notes

- The remote agent gets the same capabilities as a local client: full read access to the
  mailbox, draft creation, moves, categories, flags, saved attachments. Restrict the tool set
  on the agent side if it only needs to read.
- Keep the reverse proxy up to date; past versions had authentication bypasses affecting
  header-based access tokens.
- Rotate the three secrets independently. Changing `OUTLOOK_MCP_TOKEN` requires restarting
  the server.
- The audit log stays on the workstation (`~/.outlook-mcp/audit.log`). Remote calls are logged
  like local ones; there is no caller identity beyond the shared token.
- If a secret is pasted anywhere outside the two machines (chat, ticket, screenshot), rotate it.
