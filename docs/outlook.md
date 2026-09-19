# Outlook MCP (`microsoft-mcp-outlook`)

Local MCP server (Stdio) for **Outlook email triage** via Microsoft Graph.
The agent can search, read, draft (drafts only, **no sending**), move mails
to folders/archive/trash, label (categories) and prioritize. Deletion is
always reversible (trash, no hard delete).

## Account support matrix

| Account | Mode | Config | Status |
|---|---|---|---|
| Work (org tenant) | Delegated | `TenantId` = tenant GUID | Implemented, live test pending (only tested with a mailbox-less account so far) |
| Personal (`outlook.com` etc.) | Delegated | `TenantId` = `consumers` for a personal-only app, or `common` for an org + personal app; token v2 | ✅ Live-verified (real mails read) |
| Service / foreign mailboxes | App-Only | + `UserIdOrUpn`, `ClientSecret`, admin consent | Implemented, untested |

## 0. Which mode? (decide first, then follow exactly one recipe)

| | **A) Delegated** (recommended) | **B) App-Only** |
|---|---|---|
| For | Triaging your own mailbox | Service without user login, foreign mailboxes |
| Login | You sign in once in the browser | Runs on a secret, no login |
| Admin needed | Usually no (self-consent is often enough) | **Yes** (admin consent required) |
| Config | **2 values**: TenantId + ClientId | **4 values**: + mode, mailbox, secret |

**Rule of thumb:** Own mailbox → recipe A. Everything else → recipe B.

## 1. Prerequisites (both recipes)

- .NET 10 SDK (`dotnet --version` → 10.x)
- Microsoft 365 mailbox

## 2. Recipe A: Delegated (own mailbox, 2 values)

**Azure (one time):**

1. [Entra Admin Center](https://entra.microsoft.com) →
   **Identity → Applications → App registrations → New registration**
   - Name: e.g. `microsoft-mcp-outlook`
   - Account types: match your mailbox (single tenant is usually fine;
     for personal accounts see troubleshooting below)
   - Redirect URI: none for now
2. On the app overview page, note down the
   **Application (client) ID** and **Directory (tenant) ID**.
3. **Authentication → Add a platform → Mobile and desktop applications**
   → enable `http://localhost` (browser login).
4. **API permissions → Add → Microsoft Graph → Delegated:**
   `Mail.Read`, `Mail.ReadWrite`, `MailboxSettings.Read`, `User.Read`
   (`offline_access` is added automatically). Consent yourself or ask
   your admin, depending on the tenant.

**Config (2 values only, secrets stay outside the repo):**

```bash
dotnet user-secrets set "Graph:TenantId" "<tenant-id>" \
  --project src/MicrosoftMcp.Outlook.Host
dotnet user-secrets set "Graph:ClientId" "<client-id>" \
  --project src/MicrosoftMcp.Outlook.Host
```

Done. Everything else (`UserIdOrUpn: me`, `DelegatedFlow: Auto`) has
sane defaults. The browser opens for login on the first tool call
(the token cache kicks in afterwards).

Headless machines (no browser): additionally
`dotnet user-secrets set "Graph:DelegatedFlow" "DeviceCode"`,
then confirm the code from the server log in your browser.

## 3. Recipe B: App-Only (service, 4 values, admin required)

**Azure (one time, admin required):**

1. Steps 1–2 from recipe A (no redirect URI needed).
2. **API permissions → Add → Microsoft Graph → Application:**
   `Mail.ReadWrite` → **Grant admin consent**.
3. **Certificates & secrets → New client secret** →
   copy the value immediately (shown only once).

**Config (4 values):**

```bash
dotnet user-secrets set "Graph:AuthMode" "AppOnly" \
  --project src/MicrosoftMcp.Outlook.Host
dotnet user-secrets set "Graph:TenantId" "<tenant-id>" \
  --project src/MicrosoftMcp.Outlook.Host
dotnet user-secrets set "Graph:ClientId" "<client-id>" \
  --project src/MicrosoftMcp.Outlook.Host
dotnet user-secrets set "Graph:UserIdOrUpn" "mailbox@example.com" \
  --project src/MicrosoftMcp.Outlook.Host
dotnet user-secrets set "Graph:ClientSecret" "<secret>" \
  --project src/MicrosoftMcp.Outlook.Host
```

Instead of user-secrets you can use environment variables
(`Graph__TenantId`, `Graph__ClientId`, … – see reference below),
e.g. directly in the client config.

## 4. Run

```bash
dotnet run --project src/MicrosoftMcp.Outlook.Host
```

Stdio transport (logs to stderr, stdout stays free for JSON-RPC), config
resolved from the binary location (cwd does not matter). Startup logs
the active mode (`Graph auth mode: Delegated …`); missing required
values fail the startup immediately with a `Next:` hint.

## 5. Wire up an MCP client

VS Code (`.vscode/mcp.json`), recipe A:

```json
{
  "servers": {
    "outlook": {
      "command": "dotnet",
      "args": ["run", "--project", "/path/to/microsoft-mcp/src/MicrosoftMcp.Outlook.Host", "--no-launch-profile"],
      "env": {
        "Graph__TenantId": "<tenant-id>",
        "Graph__ClientId": "<client-id>"
      }
    }
  }
}
```

Recipe B additionally: `"Graph__AuthMode": "AppOnly"`,
`"Graph__UserIdOrUpn": "mailbox@example.com"`,
`"Graph__ClientSecret": "<secret>"`.

Claude Desktop (`claude_desktop_config.json`): same schema under
`mcpServers`. For production use take the published binary
(`dotnet publish`, or a release asset) instead of `dotnet run`.

## 6. First triage run

Example prompt for the agent:
> “Show unread mails from today (`outlook_search_emails`), summarize each in one
> sentence, archive newsletters (`outlook_archive_email`), label invoices with
> `outlook_set_categories`, and prepare reply drafts (`outlook_create_reply_draft`)
> – send nothing, delete nothing permanently.”

The server deliberately **cannot**: send (no tool, no `Mail.Send`
permission needed), hard-delete, or download attachments larger than
2 MB and nested items (see `outlook_read_attachment` limits).

## 7. Tools (17)

`outlook_list_attachments`, `outlook_read_attachment`, `outlook_list_categories` · Organize:
`outlook_move_email`, `outlook_archive_email`, `outlook_delete_email` (trash), `outlook_create_folder`,
`outlook_set_categories`, `outlook_mark_read`, `outlook_set_importance` · Draft:
`outlook_create_draft`, `outlook_create_reply_draft`, `outlook_create_forward_draft`,
`outlook_update_draft`.
`outlook_read_attachment` returns text files decoded (truncated at 20000 chars)
and binary files as base64; downloads above `maxBytes` (default 768 KB,
max 2097152) are rejected with `attachment-too-large`. Nested messages
and OneDrive links are reported, not downloaded.

Mail drafts can also be created and edited without a `policy.json`; no
administrative recipient or AI-disclosure rules apply in that case. This is
separate from the Graph permissions: `Mail.Read` is enough to read existing
messages and drafts, but `Mail.ReadWrite` is required for
`outlook_create_draft`, `outlook_create_reply_draft`,
`outlook_create_forward_draft`, and `outlook_update_draft` because they write
to the mailbox. `Mail.Send` is not required because this server does not send
messages.


`message-not-found`, `folder-not-found`, `mailbox-unavailable`, `invalid-request`,
`attachment-too-large`,
`auth-misconfigured`, `auth-failed`, `access-denied`, `throttled`,

## 9. Troubleshooting
- Cannot switch account type: open the app **Manifest** and **first**
  (Background: multi-tenant/personal audience requires access-token
  version 2; the new Authentication UI reports this only cryptically.)
- Personal account (`outlook.com` etc.): use `Graph:TenantId=consumers`
  with a personal-only app (`signInAudience=PersonalMicrosoftAccount`). Use

  `Graph:TenantId=common` only with an app that supports both org and personal
  accounts (`signInAudience=AzureADandPersonalMicrosoftAccount`). Confirm the
  device flow with the mailbox account.

- `AADSTS2346` → the app is personal-only but the authority is `common`;
  set `Graph:TenantId=consumers`, or change the app audience to
  `AzureADandPersonalMicrosoftAccount` when organizational accounts are also
  required.

- `[auth-failed]` headless → `Graph__DelegatedFlow=DeviceCode`
- `AADSTS7000218` (client_assertion/client_secret required) → in the
  app registration go to **Authentication → Allow public client flows
  → Yes** (device code / browser are public-client flows)
- `[access-denied]` → check consent/scopes from section 2
  (delegated: `Mail.ReadWrite`; app-only: `Mail.ReadWrite` + admin consent)
- `[throttled]` → wait ~60 s, smaller `top`
- Empty tool list in the client → check `dotnet build`, then inspect the
  server stderr (in the client log) for `OptionsValidationException`

## 10. Enterprise-Policy: nur interne Drafts + KI-Hinweis (`policy.json`)

Drafts an externe Adressen werden mit `[invalid-request]` abgelehnt, und jeder
Draft-Body bekommt server-seitig (nicht vom LLM, daher nicht weglassbar) den
Hinweis-Text angehängt. Kalender-Einladungen prüfen ihre eigene
Teilnehmer-Policy; der KI-Hinweis gilt dort nicht. Alles steuert die eine
**admin-owned `policy.json`** (Format: `docs/policy.example.json`, Deployment
per Skript — siehe 10.1):

```jsonc
{
  "version": 1,
  "outlook": {
    "requireInternalRecipients": true,
    "allowedRecipientDomains": ["firma.de"],
    "allowedRecipientAddresses": ["partner@firma.example"],
    "aiDisclosureEnabled": true,
    "aiDisclosureText": "Hinweis: Dieser Entwurf wurde von einer KI erstellt und muss vor dem Versand geprüft werden."
  },
  "calendar": {
    "requireInternalAttendees": true,
    "allowedAttendeeDomains": ["firma.de"],
    "allowedAttendeeAddresses": ["partner@firma.example"]
  },
  "teams": {}
}
```

Regeln: Subdomains sind eingeschlossen (`mail.firma.de` passt zu `firma.de`),
Groß-/Kleinschreibung egal, genau ein `@` erforderlich. Eine exakte Adresse in
`allowedRecipientAddresses` ist zusätzlich erlaubt, auch wenn ihre Domain nicht
in `allowedRecipientDomains` steht. Beide Listen werden als OR-Allowlist
behandelt. Die Outlook-Prüfung gilt für Mail-Empfänger; Calendar verwendet die
entsprechenden Attendee-Felder.
Antworten prüfen
server-seitig das tatsächliche Reply-Ziel (`Reply-To`, sonst Absender) und
scheitern geschlossen ohne Absender; Weiterleitungen prüfen die expliziten
Empfänger; `outlook_update_draft` ohne `to` prüft die bestehenden Empfänger des
Drafts, ohne `body` wird der Hinweis auf den bestehenden Body nachgetragen. Der
Hinweis wird idempotent angehängt (kein Doppel bei `outlook_update_draft`).

**Wichtig:** Diese Datei ist die *einzige* Quelle — `Messaging__*`-Env-Vars
werden absichtlich ignoriert, weil `mcp.json` user-schreibbar ist und ein LLM
mit Dateizugriff die Policy sonst per Env-Override aushebeln könnte.

### 10.1 Deployment per Skript (empfohlen)

Die Skripte unter `deploy/` nehmen alle Werte als Parameter entgegen, bauen das
JSON korrekt auf und schützen die Datei sofort schreibgeschützt — als Admin
ausführen:

```powershell
# Windows (Admin-PowerShell)
.\deploy\deploy-policy.ps1 -AllowedRecipientDomains firma.de,tochter.firma.de `
  -AiDisclosureText "Hinweis: Dieser Entwurf wurde von einer KI erstellt und muss vor dem Versand geprüft werden."
```

```bash
# Linux / macOS (root via sudo, Zielpfad wird je OS gewählt)
sudo ./deploy/deploy-policy.sh --domains firma.de,tochter.firma.de \
  --disclosure-text "Hinweis: Dieser Entwurf wurde von einer KI erstellt und muss vor dem Versand geprüft werden."
```

Calendar kann unabhängig konfiguriert werden. Beim Shell-Skript stehen dafür
`--calendar-domains`, `--calendar-addresses`, `--calendar-restrict` und
`--calendar-no-restrict` zur Verfügung; PowerShell verwendet
`-AllowedAttendeeDomains`, `-AllowedAttendeeAddresses` und
`-RequireInternalAttendees:$false`.

Beide Skripte validieren die Eingaben (keine Restriktion ohne Domains, kein
Hinweis ohne Text), schreiben das v1-Dokument atomar, legen es admin-owned +
read-only ab und verifizieren den Schutz (PS: ACL-Audit auf Users-Schreibrechte;
sh: `test -w` als aufrufender User). Danach Binary ebenfalls schützen (sonst
wird der Code statt der Config gepatcht).

### 10.2 Manuelle Ablage + Schutz (Fallback ohne Skript)

Suchreihenfolge: Systempfad zuerst, dann `policy.json` neben dem Binary.

| OS | Systempfad | Schutz setzen |
|---|---|---|
| Linux | `/etc/microsoft-mcp/policy.json` | `sudo install -o root -g root -m 644 policy.json /etc/microsoft-mcp/policy.json` |
| macOS | `/Library/Application Support/microsoft-mcp/policy.json` | `sudo install -o root -g wheel -m 644 policy.json "/Library/Application Support/microsoft-mcp/policy.json"` |
| Windows (Admin-PS) | `%ProgramData%\microsoft-mcp\policy.json` | Kopieren, dann `icacls policy.json /inheritance:r /grant:r Administrators:F SYSTEM:F /grant:r Users:R` |

Startverhalten: keine `policy.json` → uneingeschränkt + Warnung auf
stderr; vorhandene, aber unlesbare Datei → Start verweigert (kein Fallback auf
schwächere Policy); restriktive Datei oder restriktives Verzeichnis mit
User-Schreibrecht → **Start verweigert**
(Fail-Closed, außer der Prozess läuft elevated = Admin-Testszenario); gefunden →
Pfad + SHA-256-Präfix + wirksame Flags landen auf stderr (Audit). Ohne
Admin-Rechte auf dem Gerät (gemanagte Clients) kann ein LLM die Datei weder
ändern noch löschen.

Optionaler Backstop (gilt auch bei umgangener lokaler Config): Exchange Admin
Center → Mail flow → Rules → Disclaimer-Regel für Mails mit KI-Vermerk, plus
ggf. Extern-Sperre für das Service-Postfach. Transportregeln greifen allerdings
erst beim Senden, nicht auf Entwürfen — die lokale Injektion bleibt die
Draft-Vorschau.

### 10.3 Migration nach einem Update

Ein neues Binary liest das flache Legacy-Format weiterhin kompatibel im
Speicher. Die persistente Aufteilung wird nach dem Binary-Update explizit
ausgeführt:

```bash
microsoft-mcp policy migrate --json --write
microsoft-mcp doctor --json
```

Die Migration erstellt ein Backup und ersetzt die Datei atomar. Sie benötigt
Administrator-/Root-Rechte; ein Agent darf diese Rechte nicht selbst
beschaffen. Ohne Schreibrechte bleibt die Legacy-Datei gültig und kann später
erneut migriert werden.

## Appendix: settings reference

Precedence: user-secrets / env override `appsettings.json`.
**Never commit secrets to `appsettings.json`.**

| Setting | Env var | Default | Purpose |
|---|---|---|---|
| `Graph:AuthMode` | `Graph__AuthMode` | `Delegated` | `Delegated` or `AppOnly` |
| `Graph:TenantId` | `Graph__TenantId` | – (required) | From section 2 |
| `Graph:ClientId` | `Graph__ClientId` | – (required) | From section 2 |
| `Graph:UserIdOrUpn` | `Graph__UserIdOrUpn` | `me` | Recipe B: mailbox UPN (required) |
| `Graph:DelegatedFlow` | `Graph__DelegatedFlow` | `Auto` | Recipe A: `Auto` (= browser), `InteractiveBrowser`, `DeviceCode` (headless) |
| `Graph:DelegatedScopes` | `Graph__DelegatedScopes` | (host template) | Outlook template: `Mail.Read,Mail.ReadWrite`. Only change for special tenants |
| `Graph:ClientSecret` | `Graph__ClientSecret` | – | Recipe B only |
| `Graph:AppCredential` | `Graph__AppCredential` | `ClientSecret` | `ClientSecret`, `ManagedIdentity` (`Certificate`: not wired yet) |
