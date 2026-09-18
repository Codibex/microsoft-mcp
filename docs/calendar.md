# Calendar MCP (`microsoft-mcp-calendar`)

Local MCP server (Stdio) for **calendar access** via Microsoft Graph.
Same pattern as the [Outlook server](outlook.md): Stdio, `isError`
results with `[code]` hints, delegated or app-only auth via the shared
`Common` library. **Read-only – no create, update, delete or move.**

## 1. Prerequisites

- .NET 10 SDK (`dotnet --version` → 10.x)
- Microsoft 365 mailbox with a calendar

## 2. Entra setup (one time)

In the **same app registration** (or a new one), add the delegated
permission (no admin needed in most tenants, otherwise ask yours):

**API permissions → Add → Microsoft Graph → Delegated:**
`Calendars.Read`.

No redirect URI is needed for device-code flow; for browser login add
the **Mobile and desktop applications** platform with `http://localhost`
(see [Outlook guide](outlook.md), recipe A steps 1–3).

## 3. Config (same 2 values as Outlook)

```bash
dotnet user-secrets set "Graph:TenantId" "<tenant-id>" \
  --project src/MicrosoftMcp.Calendar.Host
dotnet user-secrets set "Graph:ClientId" "<client-id>" \
  --project src/MicrosoftMcp.Calendar.Host
```

The host template defaults `DelegatedScopes` to `Calendars.Read`.
For supported personal-account scenarios, use `Graph:TenantId=consumers` with
a personal-only app, or `common` with an org + personal app (see
[Outlook troubleshooting](outlook.md#9-troubleshooting)). Headless:
`Graph:DelegatedFlow` = `DeviceCode`.

## 4. Run

```bash
dotnet run --project src/MicrosoftMcp.Calendar.Host
```

Startup logs the mode (`[startup] Calendar auth mode: Delegated …`) and
fails fast on missing values with a `Next:` hint.

## 5. Wire up an MCP client

```json
{
  "servers": {
    "calendar": {
      "command": "/path/to/microsoft-mcp-calendar",
      "env": {
        "Graph__TenantId": "<tenant-id>",
        "Graph__ClientId": "<client-id>"
      }
    }
  }
}
```

For production use take the published binary (or a release asset).

## 6. Tools (4, read-only)

- `calendar_list_calendars` – calendars with id, name, default flag
- `calendar_list_events` – window view (default: default calendar, now plus 7 days).
  Times in ISO format, e.g. `2026-09-14T00:00:00`
- `calendar_search_events` – subject/keyword search across calendars
- `calendar_read_event` – full event (attendees, recurrence summary, body
  truncated at 8000 chars, online meeting link)

Example prompt: “What meetings do I have tomorrow (`calendar_list_events` with
`timeMin`/`timeMax`)? Who attends the sync (`calendar_read_event`)?”

## 7. Error codes (returned as `isError` results with a `Next:` hint)

`calendar-not-found`, `event-not-found`, `invalid-request`,
`auth-misconfigured`, `auth-failed`, `access-denied`, `throttled`,
`conflict`, `service-unavailable`, `graph-error`.
Details + stack traces go to the server log (stderr) only, never to the client.

## 8. Troubleshooting

- `[access-denied]` → check the `Calendars.Read` consent/permission
  from section 2 (consent again after adding permissions)
- `[auth-failed]` headless → `Graph__DelegatedFlow=DeviceCode`
- Auth/account-type issues (personal accounts, token version, public
  client flows) → same fixes as in the
  [Outlook troubleshooting](outlook.md#9-troubleshooting)
