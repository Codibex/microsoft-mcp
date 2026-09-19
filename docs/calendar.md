# Calendar MCP (`microsoft-mcp-calendar`)

Local MCP server (Stdio) for **calendar access** via Microsoft Graph.
Same pattern as the [Outlook server](outlook.md): Stdio, `isError`
results with `[code]` hints, delegated or app-only auth via the shared
`Common` library. Create and update are delegated-only; no delete or move
operation is exposed.

## 1. Prerequisites

- .NET 10 SDK (`dotnet --version` → 10.x)
- Microsoft 365 mailbox with a calendar

## 2. Entra setup (one time)

In the **same app registration** (or a new one), add the delegated
permission (no admin needed in most tenants, otherwise ask yours):

**API permissions → Add → Microsoft Graph → Delegated:**
`Calendars.Read`, `Calendars.ReadWrite`.

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

The host template defaults `DelegatedScopes` to `Calendars.Read` and
`Calendars.ReadWrite`. Create/update require delegated auth; App-Only remains
available for read operations only.
If an admin-owned `policy.json` is present, its `calendar` policy applies to
event attendees. Domains in `allowedAttendeeDomains` are allowed including
subdomains; exact exceptions can be listed in `allowedAttendeeAddresses`.
Legacy flat policies remain compatible and are mapped to these attendee fields;
Outlook AI-disclosure settings are never applied to Calendar.
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

## 6. Tools (6)

- `calendar_list_calendars` – calendars with id, name, default flag
- `calendar_list_events` – window view (default: default calendar, now plus 7 days).
  Times in ISO format, e.g. `2026-09-14T00:00:00`
- `calendar_search_events` – subject/keyword search across calendars
- `calendar_read_event` – full event (attendees, recurrence summary, body
  truncated at 8000 chars, online meeting link)
- `calendar_create_event` – create an event with an optional plain-text body,
  location, attendees, all-day flag and reminder
- `calendar_update_event` – update only the supplied event fields; no delete or
  move operation is offered

Write tools accept ISO date/time values and normalize timed events to UTC.
All-day values must be at midnight; their supplied calendar date is preserved.
Attendee values are SMTP addresses. `calendar_update_event` requires at least
one mutable field; an empty attendee list clears the attendees.

Example prompt: “What meetings do I have tomorrow (`calendar_list_events` with
`timeMin`/`timeMax`)? Who attends the sync (`calendar_read_event`)?”

## 7. Error codes (returned as `isError` results with a `Next:` hint)

`calendar-not-found`, `event-not-found`, `invalid-request`,
`auth-misconfigured`, `auth-failed`, `access-denied`, `throttled`,
`conflict`, `service-unavailable`, `graph-error`.
Details + stack traces go to the server log (stderr) only, never to the client.

## 8. Troubleshooting

- `[access-denied]` → check the `Calendars.Read` and `Calendars.ReadWrite`
  consent/permissions from section 2 (writes require delegated auth)
- `[auth-misconfigured]` on create/update → set `Graph:AuthMode` to
  `Delegated`; App-Only cannot write calendar events
- `[auth-failed]` headless → `Graph__DelegatedFlow=DeviceCode`
- Auth/account-type issues (personal accounts, token version, public
  client flows) → same fixes as in the
  [Outlook troubleshooting](outlook.md#9-troubleshooting)
