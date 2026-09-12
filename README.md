# MicrosoftMcp Outlook

Lokaler MCP-Server (Stdio) für **Outlook-Email-Triage** über Microsoft Graph.
.NET 10 / C# 14, zentrales Paketmanagement (`Directory.Packages.props`).

Der Agent kann suchen, lesen, entwerfen (nur Drafts, **kein Senden**),
in Ordner/Archiv/Papierkorb verschieben sowie labeln (Categories) und
priorisieren. Löschen ist immer reversibel (Papierkorb, kein Hard-Delete).

## 0. Welcher Modus? (zuerst entscheiden, dann nur ein Rezept befolgen)

| | **A) Delegiert** (empfohlen) | **B) App-Only** |
|---|---|---|
| Wofür | Eigenes Postfach triagieren | Dienst ohne User-Login, fremde Postfächer |
| Login | Du meldest dich einmal im Browser an | Läuft mit Secret, kein Login |
| Admin nötig | Meist nein (Selbstzustimmung reicht oft) | **Ja** (Admin-Consent Pflicht) |
| Config | **2 Werte**: TenantId + ClientId | **4 Werte**: + Modus, Postfach, Secret |

**Faustregel:** Eigenes Postfach → Rezept A. Alles andere → Rezept B.

## 1. Voraussetzungen (für beide Rezepte)

- .NET 10 SDK (`dotnet --version` → 10.x)
- Microsoft-365-Postfach

## 2. Rezept A: Delegiert (eigenes Postfach, 2 Werte)

**Azure (einmalig):**

1. [Entra Admin Center](https://entra.microsoft.com) →
   **Identität → Anwendungen → App-Registrierungen → Neue Registrierung**
   - Name: z. B. `microsoft-mcp-outlook`
   - Kontotypen: passend zum Postfach (Einzelner Mandant reicht meist)
   - Umleitungs-URI: vorerst keine
2. Auf der App-Übersicht **Anwendungs-(Client-)ID** und
   **Verzeichnis-(Mandanten-)ID** notieren.
3. **Authentifizierung → Plattform hinzufügen → Mobile und Desktopanwendungen**
   → `http://localhost` aktivieren (Browser-Login).
4. **API-Berechtigungen → Hinzufügen → Microsoft Graph → Delegiert:**
   `Mail.Read`, `Mail.ReadWrite`, `MailboxSettings.Read`, `User.Read`
   (`offline_access` kommt automatisch). Je nach Tenant zustimmen
   (selbst oder per Admin).

**Config (nur 2 Werte, Secrets landen außerhalb des Repos):**

```bash
dotnet user-secrets set "Graph:TenantId" "<tenant-id>" \
  --project src/MicrosoftMcp.Outlook.Host
dotnet user-secrets set "Graph:ClientId" "<client-id>" \
  --project src/MicrosoftMcp.Outlook.Host
```

Fertig. Alles andere (`UserIdOrUpn: me`, `DelegatedFlow: Auto`) hat
sinnvolle Defaults. Beim ersten Tool-Aufruf öffnet sich der Browser
für den Login (danach greift der Token-Cache).

Headless-Maschinen (kein Browser): zusätzlich
`dotnet user-secrets set "Graph:DelegatedFlow" "DeviceCode"`,
Code aus dem Server-Log im Browser bestätigen.

## 3. Rezept B: App-Only (Dienst, 4 Werte, Admin nötig)

**Azure (einmalig, Admin erforderlich):**

1. Schritte 1–2 aus Rezept A (keine Redirect-URI nötig).
2. **API-Berechtigungen → Hinzufügen → Microsoft Graph → Anwendung:**
   `Mail.ReadWrite` → **Administratorzustimmung erteilen**.
3. **Zertifikate & Geheimnisse → Neuer geheimer Clientschlüssel** →
   Wert sofort notieren (wird nur einmal angezeigt).

**Config (4 Werte):**

```bash
dotnet user-secrets set "Graph:AuthMode" "AppOnly" \
  --project src/MicrosoftMcp.Outlook.Host
dotnet user-secrets set "Graph:TenantId" "<tenant-id>" \
  --project src/MicrosoftMcp.Outlook.Host
dotnet user-secrets set "Graph:ClientId" "<client-id>" \
  --project src/MicrosoftMcp.Outlook.Host
dotnet user-secrets set "Graph:UserIdOrUpn" "postfach@example.com" \
  --project src/MicrosoftMcp.Outlook.Host
dotnet user-secrets set "Graph:ClientSecret" "<secret>" \
  --project src/MicrosoftMcp.Outlook.Host
```

Alternativ zu User-Secrets gehen Umgebungsvariablen
(`Graph__TenantId`, `Graph__ClientId`, … – siehe Referenz unten),
z. B. direkt in der Client-Config (Schritt 5).

## 4. Bauen, testen, starten

```bash
dotnet build MicrosoftMcp.slnx
dotnet test MicrosoftMcp.slnx
dotnet run --project src/MicrosoftMcp.Outlook.Host
```

Der Server spricht MCP über Stdio (Logs nach stderr, stdout bleibt für
JSON-RPC frei), findet seine Config anhand des Binary-Standorts (cwd egal)
und meldet beim Start den aktiven Modus im Log
(`Graph auth mode: Delegated …`). Fehlende Pflichtwerte brechen den Start
sofort mit `Next:`-Hinweis ab statt erst beim ersten Tool-Aufruf.

## 5. In MCP-Client einbinden

VS Code (`.vscode/mcp.json`), Rezept A:

```json
{
  "servers": {
    "outlook": {
      "command": "dotnet",
      "args": ["run", "--project", "/pfad/zu/microsoft-mcp/src/MicrosoftMcp.Outlook.Host", "--no-launch-profile"],
      "env": {
        "Graph__TenantId": "<tenant-id>",
        "Graph__ClientId": "<client-id>"
      }
    }
  }
}
```

Rezept B zusätzlich: `"Graph__AuthMode": "AppOnly"`,
`"Graph__UserIdOrUpn": "postfach@example.com"`,
`"Graph__ClientSecret": "<secret>"`.

Claude Desktop (`claude_desktop_config.json`): gleiches Schema unter
`mcpServers`. Für Produktivnutzung das veröffentlichte Binary nehmen
(`dotnet publish`) statt `dotnet run`.

## 6. Erster Triage-Durchlauf

Beispiel an den Agenten:
> „Zeige ungelesene Mails von heute (`search_emails`), fasse jede in einem
> Satz zusammen, archiviere Newsletter (`archive_email`), labelle Rechnungen
> mit `set_categories`, und lege Antwort-Entwürfe an (`create_reply_draft`)
> – nichts senden, nichts endgültig löschen.“

Der Server kann bewusst **nicht**: senden (kein Tool, keine
`Mail.Send`-Berechtigung nötig), endgültig löschen, Anhänge herunterladen
(nur Metadaten via `list_attachments`).

## 7. Tools (16)

Suchen/Lesen: `search_emails`, `read_email`, `list_folders`,
`list_attachments`, `list_categories` · Organisieren: `move_email`,
`archive_email`, `delete_email` (Papierkorb), `create_folder`,
`set_categories`, `mark_read`, `set_importance` · Entwerfen:
`create_draft`, `create_reply_draft`, `create_forward_draft`,
`update_draft`.

## 8. Fehlercodes (kommen als `isError`-Resultat mit `Next:`-Hinweis)

`message-not-found`, `folder-not-found`, `invalid-request`,
`auth-misconfigured`, `auth-failed`, `access-denied`, `throttled`,
`conflict`, `service-unavailable`, `graph-error`.
Details + Stacktrace landen nur im Server-Log (stderr), nie beim Client.

## 9. Troubleshooting

- `[auth-failed]` headless → `Graph__DelegatedFlow=DeviceCode`
- Token-Cache-Fehler (Linux ohne Keyring) →
  `Graph__EnableTokenCache=false`
- `[access-denied]` → Consent/Scopes aus Schritt 2 prüfen
  (delegiert: `Mail.ReadWrite`; App-Only: `Mail.ReadWrite` + Admin-Consent)
- `[throttled]` → ~60 s warten, kleineres `top`
- Leere Tool-Liste im Client → `dotnet build` prüfen, dann Server-Stderr
  (im Client-Log) auf `OptionsValidationException` kontrollieren

## Versionierung (SemVer)

`MAJOR.MINOR.PATCH`, zentral in `Directory.Build.props` (`VersionPrefix`).
Solange `MAJOR = 0` darf `MINOR` Breaking Changes enthalten.

- Neues Tool / neues Feature → `MINOR` hoch (`0.1.0` → `0.2.0`)
- Fix ohne API-Änderung → `PATCH` hoch
- Breaking Change ab `1.0.0` → `MAJOR` hoch
- Release: Tag `vX.Y.Z` pushen → GitHub Action baut, testet und
  erstellt das Release automatisch (`vX.Y.Z-rc.1` wird als Prerelease
  markiert)
- Prerelease lokal bauen: `dotnet build -p:VersionSuffix=rc.1`

## Anhang: Referenz aller Einstellungen

Rangfolge: User-Secrets / Env schlagen `appsettings.json`.
**Keine Secrets in `appsettings.json` committen.**

| Einstellung | Env-Var | Default | Wozu |
|---|---|---|---|
| `Graph:AuthMode` | `Graph__AuthMode` | `Delegated` | `Delegated` oder `AppOnly` |
| `Graph:TenantId` | `Graph__TenantId` | – (Pflicht) | Aus Schritt 2 |
| `Graph:ClientId` | `Graph__ClientId` | – (Pflicht) | Aus Schritt 2 |
| `Graph:UserIdOrUpn` | `Graph__UserIdOrUpn` | `me` | Rezept B: UPN des Postfachs (Pflicht) |
| `Graph:DelegatedFlow` | `Graph__DelegatedFlow` | `Auto` | Rezept A: `Auto` (= Browser), `InteractiveBrowser`, `DeviceCode` (headless) |
| `Graph:DelegatedScopes` | `Graph__DelegatedScopes` | `Mail.Read,Mail.ReadWrite` | Nur ändern bei Sonder-Tenants |
| `Graph:ClientSecret` | `Graph__ClientSecret` | – | Nur Rezept B |
| `Graph:AppCredential` | `Graph__AppCredential` | `ClientSecret` | `ClientSecret`, `ManagedIdentity` (`Certificate`: noch nicht verdrahtet) |
| `Graph:EnableTokenCache` | `Graph__EnableTokenCache` | `true` | Bei Keyring-Problemen auf `false` |
