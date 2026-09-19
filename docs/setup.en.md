# Consumer setup (use `microsoft-mcp`, don't develop it)

Goal: a running MCP server in your client with exactly one Entra app.
Source of truth for agents and the CLI: this file
(same content as `setup.de.md`). Domain details live in `outlook.md`,
`onedrive.md`, `calendar.md`, `teams.md`.

## 0. Three questions first (don't guess)

1. **Account:** work (org tenant, GUID) or personal (`outlook.com` etc.)?
  For a personal-only app use `TenantId=consumers`; for an app supporting
  both org and personal accounts use `TenantId=common`.
2. **Domains:** which of `outlook,onedrive,calendar,teams`?
   Unified host: `microsoft-mcp --servers outlook,calendar` (or `MCP_SERVERS`),
   no argument means all.
3. **Client:** see table — file, key and snippet format depend on
   the client (`--client`):

   | Client | `--client` | File | Key | Format |
   |---|---|---|---|---|
   | VS Code | `vscode` | `.vscode/mcp.json` | `servers` | JSON |
   | Claude Desktop | `claude` | `claude_desktop_config.json` | `mcpServers` | JSON |
   | OpenCode | `opencode` | `opencode.json` | `mcp` | JSON (`command` as array, env is called `environment`, plus `"type": "local"`) |
   | Codex CLI | `codex` | `~/.codex/config.toml` | `mcp_servers` | TOML (alternative: `codex mcp add …`) |
   | OpenClaw | `openclaw` | `openclaw.json` (`mcp.servers`) | `mcp` → `servers` | JSON (plus `"enabled": true`) |
   | Hermes Agent | `hermes` | `~/.hermes/config.yaml` | `mcp_servers` | YAML (keep secrets in `~/.hermes/.env`) |

Rule: own mailbox → Delegated. Service/foreign mailboxes → App-Only
(`outlook` only, admin consent required). `onedrive`, `calendar`, `teams`
need Delegated (`/me/*`).

## 1. Binary (no SDK needed)

Take the GitHub release assets (linux-x64/arm64, win-x64/arm64,
osx-arm64) or `dotnet publish` locally. `dotnet run` is for
development only. Remember the path, e.g. `/opt/microsoft-mcp/microsoft-mcp`.

Check:

```bash
microsoft-mcp doctor --servers outlook,calendar
```

## 2. Entra app (one time)

1. Entra Admin Center → App registrations → New registration
   (name e.g. `microsoft-mcp`, account type matching your account).
2. Overview → note down **Application (client) ID** + **Directory (tenant) ID**.
  Personal-only: `TenantId` = `consumers`; org + personal app:
  `TenantId` = `common`.
3. Authentication → Add a platform → Mobile and desktop applications →
   enable `http://localhost` (browser login; device code doesn't need it).
   Allow public client flows (otherwise `AADSTS7000218`).
4. API permissions → Add → Microsoft Graph → Delegated, exactly the scopes
   of the selected domains (least privilege, `offline_access` is automatic):

   | Domain | Scopes |
   |---|---|
   | `outlook` | `Mail.Read`, `Mail.ReadWrite` |
   | `onedrive` | `Files.Read`, `Files.ReadWrite` |
  | `calendar` | `Calendars.Read`, `Calendars.ReadWrite` |
  | `teams` | `Team.ReadBasic.All`, `ChannelMessage.Read.All`, `Chat.ReadBasic`, `Chat.Read`, `OnlineMeetingTranscript.Read.All`, `OnlineMeetingAiInsight.Read.All` |

   Then consent (yourself or your admin). App-Only (Outlook only):
   application permission `Mail.ReadWrite` + admin consent + client secret.

Personal-only apps use token version 2 + `PersonalMicrosoftAccount` and
`TenantId=consumers`. Apps supporting org + personal accounts use token
version 2 + `AzureADandPersonalMicrosoftAccount` and `TenantId=common`
(see `outlook.md` troubleshooting). Headless without browser:
`Graph__DelegatedFlow=DeviceCode`.

Token caching is secure by default: the server prefers the encrypted OS cache
and falls back to an in-memory cache when Linux Secret Service is unavailable.
The fallback keeps the process usable, but requires a new login after restart.
With persistent caching enabled, the server also stores the non-secret Azure
Identity authentication record needed to select the cached account after a
restart. It is stored in the current user's local application data directory:
normally `%LOCALAPPDATA%\\microsoft-mcp\\authentication-record.json` on
Windows and `~/.local/share/microsoft-mcp/authentication-record.json` on Linux.
The Windows token cache is protected by DPAPI; on Linux it uses Secret Service.
For persistence on a headless Linux host, install and run GNOME Keyring/libsecret
with a D-Bus session. To disable persistence explicitly, set
`Graph__EnableTokenCache=false`. Only if unencrypted disk storage is acceptable,
set `Graph__UnsafeAllowUnencryptedTokenCache=true`; this is not recommended.
When that option is enabled, the server writes a warning to stderr at startup.
If `Graph__FallbackToMemoryTokenCache=false` and the cache cannot be opened, the
server returns `auth-cache-unavailable` with the available recovery options.

## 3. Wire up the client

Secrets go into `env`, never into the repo. Generate a template:

```bash
microsoft-mcp setup --servers outlook,calendar --account personal --client vscode
microsoft-mcp setup --servers outlook --account work --client codex --binary /opt/microsoft-mcp/microsoft-mcp
```

VS Code (`.vscode/mcp.json`):

```json
{
  "servers": {
    "m365": {
      "command": "/opt/microsoft-mcp/microsoft-mcp",
      "args": ["--servers", "outlook,calendar"],
      "env": { "Graph__TenantId": "<id>", "Graph__ClientId": "<id>" }
    }
  }
}
```

Claude Desktop: same object under `mcpServers`. OpenCode/OpenClaw:
same object under `mcp` (OpenCode: `command` as array, env is called
`environment`). Codex (`~/.codex/config.toml`, TOML):

```toml
[mcp_servers.m365]
command = "/opt/microsoft-mcp/microsoft-mcp"
args = ["--servers", "outlook,calendar"]

[mcp_servers.m365.env]
Graph__TenantId = "<id>"
Graph__ClientId = "<id>"
```

Alternative: `codex mcp add m365 --env Graph__TenantId=<id> --env
Graph__ClientId=<id> -- /opt/microsoft-mcp/microsoft-mcp --servers
outlook,calendar`. Hermes (`~/.hermes/config.yaml`, YAML):

```yaml
mcp_servers:
  m365:
    command: "/opt/microsoft-mcp/microsoft-mcp"
    args: ["--servers", "outlook,calendar"]
    env:
      Graph__TenantId: "<id>"
      Graph__ClientId: "<id>"
```

App-Only additionally: `Graph__AuthMode=AppOnly`, `Graph__UserIdOrUpn`,
`Graph__ClientSecret`. Versioned `policy.json` (Outlook/Calendar/Teams,
optional) lives in the admin-owned system path or next to the binary;
`Messaging__*` env is ignored (see `outlook.md` §10).

## 4. Verify

```bash
microsoft-mcp doctor --servers outlook,calendar
microsoft-mcp doctor --json
```

Expectation: all checks `ok`, exit 0. Then in the client, e.g.:
"What is on my calendar tomorrow (`calendar_list_events`)?"
Errors come back as `isError` with a `[code]` + `Next:` hint; details only
on stderr. `[access-denied]` → consent/scopes from §2, `[auth-failed]`
headless → `DeviceCode`, empty tool list → check stderr for
`OptionsValidationException`. OpenClaw: additionally
`openclaw mcp doctor m365 --probe` for a live connection test.

### Policy migration after an update

The new binary keeps reading the old flat policy in memory, so an update does
not weaken or disable existing restrictions. To persist the split document,
run the new binary after replacing it and before restarting the MCP client:

```bash
microsoft-mcp policy migrate --json
microsoft-mcp policy migrate --json --write
microsoft-mcp doctor --json
```

The first command is a preview. `--write` is explicit, creates a backup and
replaces the admin-owned file atomically. It must run as administrator/root;
an agent such as Hermes can execute the sequence only when its update process
already has those rights. Without them, leave the legacy file in place and
ask an administrator to run the migration. The MCP client must be restarted
after the binary or policy file changes.

## Machine contract (for `setup`/`doctor`/`policy migrate`)

- Inputs: `--servers` (subset of `outlook,onedrive,calendar,teams`,
  empty = all), `--account work|personal`, `--auth delegated|apponly`,
  `--client vscode|claude|opencode|codex|openclaw|hermes|generic`,
  `--binary PATH`, `--headless`, `--json`.
- `setup` needs no secrets, writes no secrets, exit 0 for a
  valid combination, 2 for an invalid one (e.g. `teams` + `apponly`,
  `onedrive`/`calendar`/`teams` + `apponly`).
- `doctor` reads the same config as the host
  (`appsettings.json` next to the binary + env + user-secrets), checks offline:
  `TenantId` format (GUID or `common|consumers|organizations`),
  `ClientId` GUID, AuthMode vs. domains, scopes vs. selection,
  `DelegatedFlow`, `policy.json` found/protected (warning only if missing).
  No network, no token. `--json` → `{ "checks": [{ "id", "ok", "message", "next", "migrationRequired" }] }`.
- `policy migrate --json` only inspects by default. Add `--write` for the
  explicit atomic migration. The result reports `migrationRequired`,
  `written`, `backupPath` and `targetVersion`; no secrets are read or written.

Cache settings: `Graph__FallbackToMemoryTokenCache` defaults to `true` and keeps
delegated auth usable without a keyring. Persistent credential token operations
are bounded by `Graph__TokenCacheTimeoutSeconds` (default `10`, allowed range
`1`-`120`). This bounds the persistent credential's token acquisition call,
which may include silent authentication, token refresh, network access, and
Secret Service operations. It does not limit the time you may take to complete
a Device-Code login. Set the fallback to `false` only when persistent encrypted
storage is required; on Linux, `doctor` then checks whether
`org.freedesktop.secrets` is reachable and fails when it is not.
Changing the cache name can require one new login.

Linux persistent-cache prerequisite: a user-session D-Bus and a Secret Service
provider exposing `org.freedesktop.secrets` must be available, usually GNOME
Keyring/libsecret. A system-wide `dbus.service` alone is not sufficient. For
headless or client-launched sessions, the server process must inherit the user
session's `DBUS_SESSION_BUS_ADDRESS`; `XDG_RUNTIME_DIR` is recommended because
some providers use it to resolve their socket. These variables are inherited
when the client spawns the MCP child. After adding or changing them, restart
the client process that owns the MCP child, otherwise the child keeps the old
environment. Hermes may start or provide part of this environment, but it is
not a Hermes-specific server requirement. With the default memory fallback
enabled, an unavailable or unresponsive Secret Service does not prevent
delegated authentication. Once the credential switches to the memory cache, it
stays there for the lifetime of that process; restart the server to try
persistent storage again after the provider recovers.
