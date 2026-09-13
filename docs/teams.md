# Teams MCP (`microsoft-mcp-teams`)

Local MCP server (Stdio) for **read-only Teams access** via Microsoft Graph.
Same pattern as the [Outlook server](outlook.md): Stdio, `isError`
results with `[code]` hints, delegated auth via the shared `Common`
library. **Read-only – no send, no create/update/delete.**

## 1. Prerequisites

- .NET 10 SDK (`dotnet --version` → 10.x)
- Microsoft Teams membership

## 2. Entra setup (one time)

In the **same app registration** (or a new one), add the delegated
permissions (no admin needed in most tenants, otherwise ask yours):

**API permissions → Add → Microsoft Graph → Delegated:**
`Team.ReadBasic.All`, `ChannelMessage.Read.All`, `Chat.ReadBasic`, `Chat.Read`.

No redirect URI is needed for device-code flow; for browser login add
the **Mobile and desktop applications** platform with `http://localhost`
(see [Outlook guide](outlook.md), recipe A steps 1–3).

App-Only is **not** supported by this host: Teamwork APIs run via
`/me/*` and need a signed-in user (service access would require a
different permission/path scheme – out of scope).

## 3. Config (same 2 values as Outlook)

```bash
dotnet user-secrets set "Graph:TenantId" "<tenant-id>" \
  --project src/MicrosoftMcp.Teams.Host
dotnet user-secrets set "Graph:ClientId" "<client-id>" \
  --project src/MicrosoftMcp.Teams.Host
```

The host template defaults `DelegatedScopes` to the Teams scopes.
Personal accounts: `Graph:TenantId` = `common`
(see [Outlook troubleshooting](outlook.md#9-troubleshooting)). Headless:
`Graph:DelegatedFlow` = `DeviceCode`.

## 4. Run

```bash
dotnet run --project src/MicrosoftMcp.Teams.Host
```

Startup logs the mode (`[startup] Teams auth mode: Delegated …`) and
fails fast on missing values with a `Next:` hint.

## 5. Wire up an MCP client

```json
{
  "servers": {
    "teams": {
      "command": "/path/to/microsoft-mcp-teams",
      "env": {
        "Graph__TenantId": "<tenant-id>",
        "Graph__ClientId": "<client-id>"
      }
    }
  }
}
```

For production use take the published binary (or a release asset).

## 6. Tools (9, read-only)

- `teams_list_teams` – joined teams with id and name
- `teams_list_channels` – channels of a team
- `teams_list_channel_messages` – channel messages (pass `teamId` + `channelId`)
- `teams_list_message_replies` – thread replies to a channel message
- `teams_read_channel_message` – full message (HTML content truncated at
  8000 chars, reactions, mentions)
- `teams_list_chats` – recent 1:1 and group chats
- `teams_list_chat_messages` – messages of a chat
- `teams_list_chat_replies` – thread replies to a chat message
- `teams_read_chat_message` – full chat message

Message previews are plain text (HTML stripped, truncated at 500 chars).
Example prompt: “What happened in #general today
(`teams_list_channel_messages`)? Show the thread (`teams_list_message_replies`).”

## 7. Error codes (returned as `isError` results with a `Next:` hint)

`team-not-found`, `channel-not-found`, `chat-not-found`,
`channel-message-not-found`, `invalid-request`,
`auth-misconfigured`, `auth-failed`, `access-denied`, `throttled`,
`conflict`, `service-unavailable`, `graph-error`.
Details + stack traces go to the server log (stderr) only, never to the client.

## 8. Troubleshooting

- `[access-denied]` → check the Teams consent/scopes from section 2
  (all four delegated scopes, consent again after adding permissions)
- `[auth-failed]` headless → `Graph__DelegatedFlow=DeviceCode`
- App-Only configured → switch to delegated; this host rejects
  `AuthMode: AppOnly` at startup with a hint
- Auth/account-type issues (personal accounts, token version, public
  client flows) → same fixes as in the
  [Outlook troubleshooting](outlook.md#9-troubleshooting)
