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
`Team.ReadBasic.All`, `ChannelMessage.Read.All`, `Chat.ReadBasic`, `Chat.Read`,
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

## 6. Tools (13, read-only)

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
  (all six delegated scopes, consent again after adding permissions). For
  insights also verify the Microsoft 365 Copilot license.
- `[auth-failed]` headless → `Graph__DelegatedFlow=DeviceCode`
- App-Only configured → switch to delegated; this host rejects
  `AuthMode: AppOnly` at startup with a hint
- Auth/account-type issues (personal accounts, token version, public
  client flows) → same fixes as in the
  [Outlook troubleshooting](outlook.md#9-troubleshooting)

## 9. Enterprise-Policy (Vorbereitung für Write)

Read-only today: there is no send path, so nothing to enforce — but the shared
admin-owned versioned `policy.json` (same file and format as
[Outlook](outlook.md#10-enterprise-policy-nur-interne-drafts--ki-hinweis-policyjson),
loaded from the OS-specific system path, `Messaging__*` env ignored) is already
validated at startup of every Teams host. Its `teams` section is intentionally
empty for now, so the file stays the single source when send tools land.

Send design (to be implemented): new tools (`teams_send_channel_message`,
`teams_send_chat_message`) plus delegated scopes `ChannelMessage.Send` and
`ChatMessage.Send` (admin consent required). Same two guarantees as mail,
enforced in server code via the shared blocks in `MicrosoftMcp.Common`:

- **Intern only:** no mail domains in Teams — instead same-tenant members.
  Before send, list members server-side and compare each
  `aadUserConversationMember.tenantId` against the configured `Graph:TenantId`
  (chat members readable with the existing `Chat.ReadBasic` scope; team/channel
  guest checks need `TeamMember.Read.All`, admin consent). Any external/guest
  member → `[invalid-request]`, fail closed.
- **Disclosure:** `MessageDisclosure.Apply(body, isHtml: true, policy)` appends
  the admin text as HTML badge — a tool parameter for it must never exist.

Note the platform limit: Teams chat messages do not pass Exchange transport
rules, so unlike mail there is no server-side disclaimer backstop — the code
guard plus scope assignment plus Purview (DLP / Communication Compliance,
detective) are the enforcement story.
