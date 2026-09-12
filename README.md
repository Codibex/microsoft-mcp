# MicrosoftMcp Outlook

Local MCP server (Stdio) for **Outlook email triage** via Microsoft Graph.
.NET 10 / C# 14, central package management (`Directory.Packages.props`).

The agent can search, read, draft (drafts only, **no sending**),
move mails to folders/archive/trash, label (categories) and prioritize.
Deletion is always reversible (trash, no hard delete).

## Account support matrix

| Account | Mode | Config | Status |
|---|---|---|---|
| Work (org tenant) | Delegated | `TenantId` = tenant GUID | Implemented, live test pending (only tested with a mailbox-less account so far) |
| Personal (`outlook.com` etc.) | Delegated | `TenantId` = `common`, app: org + personal, token v2 | ✅ Live-verified (real mails read) |
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
   - Account types: match your mailbox (single tenant is usually fine)
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
e.g. directly in the client config (step 5).

## 4. Build, test, run

```bash
dotnet build MicrosoftMcp.slnx
dotnet test MicrosoftMcp.slnx
dotnet run --project src/MicrosoftMcp.Outlook.Host
```

The server speaks MCP over stdio (logs go to stderr, stdout stays free
for JSON-RPC), resolves its config from the binary location (cwd does
not matter) and logs the active mode on startup
(`Graph auth mode: Delegated …`). Missing required values fail the
startup immediately with a `Next:` hint instead of the first tool call.

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
> “Show unread mails from today (`search_emails`), summarize each in one
> sentence, archive newsletters (`archive_email`), label invoices with
> `set_categories`, and prepare reply drafts (`create_reply_draft`)
> – send nothing, delete nothing permanently.”

The server deliberately **cannot**: send (no tool, no `Mail.Send`
permission needed), hard-delete, or download attachments larger than
2 MB and nested items (see `read_attachment` limits).

## 7. Tools (17)

Search/read: `search_emails`, `read_email`, `list_folders`,
`list_attachments`, `read_attachment`, `list_categories` · Organize: `move_email`,
`archive_email`, `delete_email` (trash), `create_folder`,
`set_categories`, `mark_read`, `set_importance` · Draft:
`create_draft`, `create_reply_draft`, `create_forward_draft`,
`update_draft`.

`read_attachment` returns text files decoded (truncated at 20000 chars)
and binary files as base64; downloads above `maxBytes` (default 768 KB,
max 2097152) are rejected with `attachment-too-large`. Nested messages
and OneDrive links are reported, not downloaded.

## 8. Error codes (returned as `isError` results with a `Next:` hint)

`message-not-found`, `folder-not-found`, `mailbox-unavailable`, `invalid-request`,
`attachment-too-large`,
`auth-misconfigured`, `auth-failed`, `access-denied`, `throttled`,
`conflict`, `service-unavailable`, `graph-error`.
Details + stack traces go to the server log (stderr) only, never to the client.

## 9. Troubleshooting

- Cannot switch account type: open the app **Manifest** and **first**
  set `"requestedAccessTokenVersion"` to `2`, save, **then** set
  `"signInAudience"` to `"AzureADandPersonalMicrosoftAccount"` and save.
  (Background: multi-tenant/personal audience requires access-token
  version 2; the new Authentication UI reports this only cryptically.)
  Recommended: org + personal, not “personal only”.
- Personal account (`outlook.com` etc.): set `Graph:TenantId` to `common`
  (or `consumers`) instead of the tenant GUID and confirm the device
  flow with the mailbox account.

- `[auth-failed]` headless → `Graph__DelegatedFlow=DeviceCode`
- `AADSTS7000218` (client_assertion/client_secret required) → in the
  app registration go to **Authentication → Allow public client flows
  → Yes** (device code / browser are public-client flows)
- Token cache errors (Linux without keyring) →
  `Graph__EnableTokenCache=false`
- `[access-denied]` → check consent/scopes from step 2
  (delegated: `Mail.ReadWrite`; app-only: `Mail.ReadWrite` + admin consent)
- `[throttled]` → wait ~60 s, smaller `top`
- Empty tool list in the client → check `dotnet build`, then inspect the
  server stderr (in the client log) for `OptionsValidationException`

## Versioning (SemVer)

`MAJOR.MINOR.PATCH`, centralized in `Directory.Build.props` (`VersionPrefix`).
While `MAJOR = 0`, `MINOR` may contain breaking changes.

- New tool / feature → bump `MINOR` (`0.1.0` → `0.2.0`)
- Fix without API change → bump `PATCH`
- Breaking change from `1.0.0` on → bump `MAJOR`
- Release: push tag `vX.Y.Z` → GitHub Action builds, tests and creates
  the release automatically (`vX.Y.Z-rc.1` is marked as prerelease) –
  including ready-made binaries (linux-x64, win-x64, osx-arm64) and
  SHA256 checksums as assets
- Build a prerelease locally: `dotnet build -p:VersionSuffix=rc.1`

## License

MIT – see [LICENSE](LICENSE).

## Appendix: settings reference

Precedence: user-secrets / env override `appsettings.json`.
**Never commit secrets to `appsettings.json`.**

| Setting | Env var | Default | Purpose |
|---|---|---|---|
| `Graph:AuthMode` | `Graph__AuthMode` | `Delegated` | `Delegated` or `AppOnly` |
| `Graph:TenantId` | `Graph__TenantId` | – (required) | From step 2 |
| `Graph:ClientId` | `Graph__ClientId` | – (required) | From step 2 |
| `Graph:UserIdOrUpn` | `Graph__UserIdOrUpn` | `me` | Recipe B: mailbox UPN (required) |
| `Graph:DelegatedFlow` | `Graph__DelegatedFlow` | `Auto` | Recipe A: `Auto` (= browser), `InteractiveBrowser`, `DeviceCode` (headless) |
| `Graph:DelegatedScopes` | `Graph__DelegatedScopes` | `Mail.Read,Mail.ReadWrite` | Only change for special tenants |
| `Graph:ClientSecret` | `Graph__ClientSecret` | – | Recipe B only |
| `Graph:AppCredential` | `Graph__AppCredential` | `ClientSecret` | `ClientSecret`, `ManagedIdentity` (`Certificate`: not wired yet) |
| `Graph:EnableTokenCache` | `Graph__EnableTokenCache` | `true` | Set to `false` on keyring issues |
