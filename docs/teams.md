# Teams MCP (`microsoft-mcp-teams`)

Local MCP server (Stdio) for Teams access via Microsoft Graph. Same pattern as
the [Outlook server](outlook.md): Stdio, `isError` results with `[code]` hints,
delegated auth via the shared `Common` library. Message sends are limited to the
two guarded send tools described below; no other write operations are exposed.

## 1. Prerequisites

- .NET 10 SDK (`dotnet --version` → 10.x)
- Microsoft Teams membership

## 2. Entra setup (one time)

In the **same app registration** (or a new one), add the delegated
permissions (no admin needed in most tenants, otherwise ask yours):

**API permissions → Add → Microsoft Graph → Delegated:**
`User.Read`, `Team.ReadBasic.All`, `ChannelMessage.Read.All`, `ChannelMember.Read.All`,
`Channel.ReadBasic.All`, `ChannelMessage.Send`, `Chat.ReadBasic`, `Chat.Read`, `ChatMessage.Send`,
`OnlineMeetingTranscript.Read.All`, `OnlineMeetingAiInsight.Read.All`.

`OnlineMeetingAiInsight.Read.All` accesses the Meeting AI Insights API. The
signed-in user must have a Microsoft 365 Copilot license. The transcript and
insight APIs are for work/school accounts; personal Microsoft accounts are not
supported. Transcription must be enabled for the meeting, and the tenant must
allow Graph transcript access.

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

The host template defaults `DelegatedScopes` to the Teams, transcript and
Meeting AI Insights scopes.
For supported personal-account scenarios, use `Graph:TenantId=consumers` with
a personal-only app, or `common` with an org + personal app (see
[Outlook troubleshooting](outlook.md#9-troubleshooting)). Headless:
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

## 6. Tools (14)

- `teams_list_teams` – joined teams with id and name
- `teams_list_channels` – channels of a team
- `teams_list_channel_messages` – channel messages (pass `teamId` + `channelId`)
- `teams_list_message_replies` – thread replies to a channel message
- `teams_read_channel_message` – full message (HTML content truncated at
  8000 chars, reactions, mentions)
- `teams_send_channel_message` – send a guarded message to an existing channel
- `teams_list_chats` – recent 1:1 and group chats
- `teams_list_chat_messages` – messages of a chat
- `teams_read_chat_message` – full chat message
- `teams_send_chat_message` – send a guarded message to an existing chat
- `teams_list_meeting_transcripts` – transcripts of a scheduled online meeting
- `teams_read_meeting_transcript` – VTT content of a transcript
- `teams_list_meeting_insights` – AI insight metadata for a completed meeting
- `teams_read_meeting_insight` – AI-generated notes, action items and mentions

Message previews are plain text (HTML stripped, truncated at 500 chars).
Example prompt: “What happened in #general today
(`teams_list_channel_messages`)? Show the thread (`teams_list_message_replies`).”

Meeting example: “Read the transcript for meeting `<meetingId>` using
`teams_list_meeting_transcripts` and `teams_read_meeting_transcript`, then
summarize its action items from `teams_read_meeting_insight`.” Insights are
available only after the meeting and may take up to four hours to appear.

## 7. Error codes (returned as `isError` results with a `Next:` hint)

`team-not-found`, `channel-not-found`, `chat-not-found`,
`channel-message-not-found`, `meeting-transcript-not-found`,
`meeting-insight-not-found`, `invalid-request`,
`auth-misconfigured`, `auth-failed`, `access-denied`, `throttled`,
`conflict`, `service-unavailable`, `graph-error`.
Details + stack traces go to the server log (stderr) only, never to the client.

## 8. Troubleshooting
- `[access-denied]` → check the Teams/meeting consent/scopes from section 2
  (all listed delegated scopes, consent again after adding permissions). For
  insights also verify the Microsoft 365 Copilot license.
- `[auth-failed]` headless → `Graph__DelegatedFlow=DeviceCode`
- App-Only configured → switch to delegated; this host rejects
  `AuthMode: AppOnly` at startup with a hint
- Auth/account-type issues (personal accounts, token version, public
  client flows) → same fixes as in the
  [Outlook troubleshooting](outlook.md#9-troubleshooting)

## 9. Enterprise policy

The shared admin-owned versioned `policy.json` (same file and format as
[Outlook](outlook.md), with the complete property reference in
[the policy reference](outlook.md#policy-v1-schema), loaded from the OS-specific system
path, with `Messaging__*` env ignored) is validated at startup of every Teams
host. Configure the `teams` section with the same domain/exact-address and
disclosure fields as Outlook:

```json
"teams": {
  "requireInternalRecipients": true,
  "allowedRecipientDomains": ["firma.de"],
  "allowedRecipientAddresses": [],
  "aiDisclosureEnabled": true,
  "aiDisclosureText": "This message was created by AI and must be reviewed before sending."
}
```

Before either send, the server lists all conversation members across every
Graph page and verifies every member as an `aadUserConversationMember`, with a
non-empty email address allowed by `policy.json`. Channel checks include
indirect members of shared channels. A concrete `Graph:TenantId` also has to
match each member's tenant id; `common`, `consumers`, and `organizations` rely
on the configured domain/exact-address policy because they are tenant
selectors, not resource tenant ids. Guest roles, foreign tenants, unknown
member types, and missing identity data are rejected before Graph receives the
message POST.

The disclosure is appended server-side by `MessageDisclosure.Apply`; it is not a
tool parameter and cannot be omitted or changed by the caller. The send tools
are `teams_send_channel_message` and `teams_send_chat_message`, and they only
target existing Graph channel/chat ids.

Note the platform limit: Teams chat messages do not pass Exchange transport
rules, so unlike mail there is no server-side disclaimer backstop. The code
guard, admin-owned policy, scope assignment, and Purview controls (DLP /
Communication Compliance) should be used together.
