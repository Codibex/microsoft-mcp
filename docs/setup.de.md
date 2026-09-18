# Setup für Verwender (`microsoft-mcp` nutzen, nicht entwickeln)

Ziel: ein laufender MCP-Server in deinem Client mit genau einer Entra-App.
Quelle der Wahrheit für Agenten und CLI: diese Datei
(inhaltsgleich mit `setup.en.md`). Domänen-Details stehen in `outlook.md`,
`onedrive.md`, `calendar.md`, `teams.md`.

## 0. Drei Fragen zuerst (nicht raten)

1. **Konto:** Arbeit (Org-Tenant, GUID) oder Personal (`outlook.com` etc.)?
  Für eine reine Personal-App `TenantId=consumers` verwenden; für eine App
  für Organisations- und persönliche Konten `TenantId=common`.
2. **Domains:** welche von `outlook,onedrive,calendar,teams`?
   Unified host: `microsoft-mcp --servers outlook,calendar` (oder `MCP_SERVERS`),
   ohne Angabe alle.
3. **Client:** siehe Tabelle — Datei, Schlüssel und Snippet-Format hängen
   vom Client ab (`--client`):

   | Client | `--client` | Datei | Schlüssel | Format |
   |---|---|---|---|---|
   | VS Code | `vscode` | `.vscode/mcp.json` | `servers` | JSON |
   | Claude Desktop | `claude` | `claude_desktop_config.json` | `mcpServers` | JSON |
   | OpenCode | `opencode` | `opencode.json` | `mcp` | JSON (`command` als Array, Env heißt `environment`, plus `"type": "local"`) |
   | Codex CLI | `codex` | `~/.codex/config.toml` | `mcp_servers` | TOML (alternativ: `codex mcp add …`) |
   | OpenClaw | `openclaw` | `openclaw.json` (`mcp.servers`) | `mcp` → `servers` | JSON (plus `"enabled": true`) |
   | Hermes Agent | `hermes` | `~/.hermes/config.yaml` | `mcp_servers` | YAML (Secrets besser in `~/.hermes/.env`) |

Regel: eigenes Postfach → Delegated. Service/fremde Postfächer → App-Only
(nur `outlook`, Admin-Consent nötig). `onedrive`, `calendar`, `teams`
brauchen Delegated (`/me/*`).

## 1. Binary (kein SDK nötig)

Release-Assets von GitHub nehmen (linux-x64/arm64, win-x64/arm64,
osx-arm64) oder lokal `dotnet publish`. `dotnet run` ist nur für
Entwicklung. Pfad merken, z. B. `/opt/microsoft-mcp/microsoft-mcp`.

Prüfen:

```bash
microsoft-mcp doctor --servers outlook,calendar
```

## 2. Entra-App (einmalig)

1. Entra Admin Center → App registrations → New registration
   (Name z. B. `microsoft-mcp`, Account-Typ passend zum Konto).
2. Overview → **Application (client) ID** + **Directory (tenant) ID** notieren.
  Nur persönliche Konten: `TenantId` = `consumers`; Organisations- und
  persönliche Konten: `TenantId` = `common`.
3. Authentication → Add a platform → Mobile and desktop applications →
   `http://localhost` aktivieren (Browser-Login; Device-Code braucht das nicht).
   Public-Client-Flows zulassen (sonst `AADSTS7000218`).
4. API permissions → Add → Microsoft Graph → Delegated, genau die Scopes
   der gewählten Domains (Least Privilege, `offline_access` kommt automatisch):

   | Domain | Scopes |
   |---|---|
   | `outlook` | `Mail.Read`, `Mail.ReadWrite` |
   | `onedrive` | `Files.Read`, `Files.ReadWrite` |
   | `calendar` | `Calendars.Read` |
   | `teams` | `Team.ReadBasic.All`, `ChannelMessage.Read.All`, `Chat.ReadBasic`, `Chat.Read` |

   Danach consentieren (selbst oder Admin). App-Only (nur Outlook):
   Application-Permission `Mail.ReadWrite` + Admin-Consent + Client-Secret.

Reine Personal-Apps brauchen Token-Version 2 + `PersonalMicrosoftAccount`
und `TenantId=consumers`. Apps für Organisations- und persönliche Konten
brauchen Token-Version 2 + `AzureADandPersonalMicrosoftAccount` und
`TenantId=common` (siehe `outlook.md` Troubleshooting). Headless ohne Browser:
`Graph__DelegatedFlow=DeviceCode`.

Der Token-Cache ist standardmaessig sicher: Der Server bevorzugt den
verschluesselten OS-Cache und wechselt auf einen In-Memory-Cache, wenn der Linux
Secret Service nicht verfuegbar ist. Der Prozess bleibt dadurch nutzbar, nach
einem Neustart ist jedoch eine neue Anmeldung noetig. Fuer Persistenz auf einem
headless Linux-Host GNOME Keyring/libsecret mit einer D-Bus-Session einrichten.
Persistenz laesst sich mit `Graph__EnableTokenCache=false` vollstaendig
deaktivieren. Nur wenn unverschluesselte Speicherung auf der Festplatte
akzeptabel ist, `Graph__UnsafeAllowUnencryptedTokenCache=true` setzen; das wird
nicht empfohlen.
Wenn diese Option aktiviert ist, schreibt der Server beim Start eine Warnung
auf stderr. Wenn `Graph__FallbackToMemoryTokenCache=false` gesetzt ist und der
Cache nicht geöffnet werden kann, liefert der Server `auth-cache-unavailable`
mit den möglichen Wiederherstellungsoptionen.

## 3. Client verdrahten

Secrets gehören in `env`, nie ins Repo. Template generieren:

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

Claude Desktop: gleiches Objekt unter `mcpServers`. OpenCode/OpenClaw:
gleiches Objekt unter `mcp` (OpenCode: `command` als Array, Env heißt
`environment`). Codex (`~/.codex/config.toml`, TOML):

```toml
[mcp_servers.m365]
command = "/opt/microsoft-mcp/microsoft-mcp"
args = ["--servers", "outlook,calendar"]

[mcp_servers.m365.env]
Graph__TenantId = "<id>"
Graph__ClientId = "<id>"
```

Alternativ: `codex mcp add m365 --env Graph__TenantId=<id> --env
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

App-Only zusätzlich: `Graph__AuthMode=AppOnly`, `Graph__UserIdOrUpn`,
`Graph__ClientSecret`. `policy.json` (nur Outlook/Teams, optional):
admin-owned Systempfad oder neben dem Binary; `Messaging__*`-Env wird
ignoriert (siehe `outlook.md` §10).

## 4. Verifizieren

```bash
microsoft-mcp doctor --servers outlook,calendar
microsoft-mcp doctor --json
```

Erwartung: alle Checks `ok`, Exit 0. Dann im Client z. B.:
„Was steht morgen im Kalender (`calendar_list_events`)?“
Fehler kommen als `isError` mit `[code]` + `Next:`-Hinweis; Details nur
auf stderr. `[access-denied]` → Consent/Scopes aus §2, `[auth-failed]`
headless → `DeviceCode`, leere Tool-Liste → stderr auf
`OptionsValidationException` prüfen. OpenClaw: zusätzlich
`openclaw mcp doctor m365 --probe` für einen Live-Verbindungstest.

## Maschinen-Vertrag (für `setup`/`doctor`)

- Eingaben: `--servers` (Teilmenge von `outlook,onedrive,calendar,teams`,
  leer = alle), `--account work|personal`, `--auth delegated|apponly`,
  `--client vscode|claude|opencode|codex|openclaw|hermes|generic`,
  `--binary PATH`, `--headless`, `--json`.
- `setup` braucht keine Secrets, schreibt keine Secrets, Exit 0 bei
  gültiger Kombination, 2 bei ungültiger (z. B. `teams` + `apponly`,
  `onedrive`/`calendar`/`teams` + `apponly`).
- `doctor` liest dieselbe Config wie der Host
  (`appsettings.json` beim Binary + Env + User-Secrets), prüft offline:
  `TenantId`-Format (GUID oder `common|consumers|organizations`),
  `ClientId`-GUID, AuthMode vs. Domains, Scopes vs. Auswahl,
  `DelegatedFlow`, `policy.json`-Fund/Schutz (nur Warnung wenn fehlend).
  Kein Netzwerk, kein Token. `--json` → `{ "checks": [{ "id", "ok", "message", "next" }] }`.

Cache-Einstellungen: `Graph__FallbackToMemoryTokenCache` ist standardmaessig
`true` und haelt Delegated-Auth ohne Keyring nutzbar. Nur wenn persistenter,
verschluesselter Speicher zwingend erforderlich ist, auf `false` setzen; auf
Linux prueft `doctor` dann, ob `org.freedesktop.secrets` erreichbar ist, und
schlaegt fehl, wenn der Dienst nicht erreichbar ist. Eine Aenderung des
Cache-Namens kann eine einmalige neue Anmeldung erfordern.
