# OneDrive MCP (`microsoft-mcp-onedrive`)

Local MCP server (Stdio) for **OneDrive files** via Microsoft Graph.
Same pattern as the [Outlook server](outlook.md): Stdio, `isError`
results with `[code]` hints, delegated auth via the shared `Common`
library. Read, create and move only – **no delete**.

## 1. Prerequisites

- .NET 10 SDK (`dotnet --version` → 10.x)
- Microsoft 365 OneDrive
- An Entra app registration (can be the same one as for Outlook)

## 2. Entra setup (one time)

In the **same app registration** (or a new one), add the delegated
permissions (no admin needed in most tenants, otherwise ask yours):

**API permissions → Add → Microsoft Graph → Delegated:**
`Files.Read`, `Files.ReadWrite` (`offline_access` is added automatically).

No redirect URI is needed for device-code flow; for browser login add
the **Mobile and desktop applications** platform with `http://localhost`
(see [Outlook guide](outlook.md), recipe A steps 1–3).

App-Only is **not** supported by this host: OneDrive access runs via
`/me/drive` and needs a signed-in user (service access would require
`Sites.Selected` and a different path scheme – out of scope).

## 3. Config (same 2 values as Outlook)

```bash
dotnet user-secrets set "Graph:TenantId" "<tenant-id>" \
  --project src/MicrosoftMcp.OneDrive.Host
dotnet user-secrets set "Graph:ClientId" "<client-id>" \
  --project src/MicrosoftMcp.OneDrive.Host
```

The host template defaults `DelegatedScopes` to the Files scopes, so no
scope config is needed. For supported personal-account scenarios, use
`Graph:TenantId=consumers` with a personal-only app, or `common` with an org
and personal app (see [Outlook troubleshooting](outlook.md#9-troubleshooting)).
Headless:
`Graph:DelegatedFlow` = `DeviceCode`.
OneDrive has no `policy.json` section because its tools do not send to
recipients, invite attendees or append AI disclosures.

## 4. Run

```bash
dotnet run --project src/MicrosoftMcp.OneDrive.Host
```

Startup logs the mode (`[startup] OneDrive auth mode: Delegated …`) and
fails fast on missing values with a `Next:` hint.

## 5. Wire up an MCP client

```json
{
  "servers": {
    "onedrive": {
      "command": "/path/to/microsoft-mcp-onedrive",
      "env": {
        "Graph__TenantId": "<tenant-id>",
        "Graph__ClientId": "<client-id>"
      }
    }
  }
}
```

For production use take the published binary (or a release asset).

## 6. Addressing items

Every item tool accepts `"root"`, an item id (from `onedrive_list_children` or
`onedrive_search_files`), or a `/path/from/root`:

- `onedrive_list_children("/Belege")`, `onedrive_get_item("01ABC…")`, `onedrive_download_file("/Belege/rechnung.pdf")`

## 7. Tools (9)

- `onedrive_get_drive` – default drive with quota (total/used/remaining)
- `onedrive_list_drives` – accessible drives (own OneDrive, shared libraries)
- `onedrive_get_item` – metadata of one file or folder
- `onedrive_list_children` – folder contents, folders first (default `root`, up to 200)
- `onedrive_search_files` – name/keyword search across the drive
- `onedrive_download_file` – without `localPath`: text decoded (truncated at
  20000 chars), binary as base64; above `maxBytes` (default 768 KB, max 2097152)
  rejected with `attachment-too-large`. **With `localPath` (absolute host path):
  streams the bytes to disk instead** – the result carries metadata + path only,
  so large files bypass the model context. Refuses to overwrite without
  `overwrite: true`
- `onedrive_create_folder` – renames on name conflict instead of failing
- `onedrive_upload_file` – `localPath` (absolute host path, read from disk so
  bytes bypass the model context) or small inline `contentText`/`contentBase64`.
  Local files up to 4194304 bytes use simple upload, larger ones (up to
  104857600 bytes) a resumable Graph upload session (`createUploadSession`,
  same v1.0 endpoint – no endpoint switch needed). Above that: OneDrive UI
- `onedrive_move_item` – rename and/or reparent; move back to undo

Large-file transfers never pass through the model: the model only sends short
path strings (like `cp` arguments), the host process moves the bytes. Paths must
be absolute. Note: like `cp`, a path-based upload can exfiltrate any host file
the process can read – only point it at files you intend to share.

## 8. Error codes (returned as `isError` results with a `Next:` hint)

`item-not-found`, `invalid-request`, `attachment-too-large`,
`auth-misconfigured`, `auth-failed`, `access-denied`, `throttled`,
`conflict`, `service-unavailable`, `graph-error`.
Details + stack traces go to the server log (stderr) only, never to the client.

## 9. Troubleshooting

- `[access-denied]` → check the Files consent/scopes from section 2
  (delegated `Files.ReadWrite`, consent again after adding permissions)
- `[auth-failed]` headless → `Graph__DelegatedFlow=DeviceCode`
- App-Only configured → switch to delegated; this host rejects
  `AuthMode: AppOnly` at startup with a hint
- Auth/account-type issues (personal accounts, token version, public
  client flows) → same fixes as in the
  [Outlook troubleshooting](outlook.md#9-troubleshooting)
