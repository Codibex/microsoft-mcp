# SharePoint MCP (`microsoft-mcp-sharepoint`)

Local MCP server (Stdio) for **SharePoint sites and files** via Microsoft
Graph. Same pattern as the [OneDrive server](onedrive.md): Stdio, `isError`
results with `[code]` hints, delegated auth via the shared `Common`
library. Read, create and move only – **no delete**.

## Relationship to OneDrive and Teams (clean split)

- `onedrive_*` tools stay pinned to the signed-in user's default drive
  (`/me/drive`): the direct path for personal files, unchanged.
- `sharepoint_*` tools always take an **explicit `driveId`**: the second
  path for everything else – site libraries and Teams channel libraries.
- `teams_get_channel_files_folder(teamId, channelId)` (see
  [Teams guide](teams.md)) is the bridge: channel files live in the
  SharePoint document library of the team's group site, and the bridge
  resolves them to `driveId` + `folderId` for the `sharepoint_*` tools.
- Chat file attachments live in the sender's OneDrive, not in SharePoint;
  use the `onedrive_*` tools for those.

## 1. Prerequisites

- .NET 10 SDK (`dotnet --version` → 10.x)
- Microsoft 365 SharePoint access
- An Entra app registration (can be the same one as for Outlook/Teams)

## 2. Entra setup (one time)

In the **same app registration** (or a new one), add the delegated
permissions (no admin needed in most tenants, otherwise ask yours):

**API permissions → Add → Microsoft Graph → Delegated:**
`Sites.Read.All`, `Files.Read`, `Files.ReadWrite` (`offline_access` is added automatically).

For Teams channel libraries in the same setup, add the Teams scopes from
the [Teams guide](teams.md) as well (union is requested automatically).

No redirect URI is needed for device-code flow; for browser login add
the **Mobile and desktop applications** platform with `http://localhost`
(see [Outlook guide](outlook.md), recipe A steps 1–3).

App-Only is **not** supported by this host: SharePoint access runs via
`/sites/*` and needs a signed-in user (service access would require
`Sites.Selected` and a different path scheme – out of scope).

Work or school accounts only: the Sites APIs (`Sites.Read.All` delegated)
are not supported for personal Microsoft accounts, so `TenantId=consumers`
cannot work here – use your org tenant GUID.

## 3. Config (same 2 values as Outlook)

```bash
dotnet user-secrets set "Graph:TenantId" "<tenant-id>" \
  --project src/MicrosoftMcp.SharePoint.Host
dotnet user-secrets set "Graph:ClientId" "<client-id>" \
  --project src/MicrosoftMcp.SharePoint.Host
```

The host template defaults `DelegatedScopes` to the SharePoint scopes, so no
scope config is needed. Personal Microsoft accounts are not supported
(see section 2); headless setups use
`Graph:DelegatedFlow` = `DeviceCode`.
SharePoint has no `policy.json` section because its tools do not send to
recipients, invite attendees or append AI disclosures.

## 4. Run

```bash
dotnet run --project src/MicrosoftMcp.SharePoint.Host
```

Startup logs the mode (`[startup] SharePoint auth mode: Delegated …`) and
fails fast on missing values with a `Next:` hint.

## 5. Wire up an MCP client

```json
{
  "servers": {
    "sharepoint": {
      "command": "/path/to/microsoft-mcp-sharepoint",
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

Site discovery first, then files – always with an explicit `driveId`:

- `sharepoint_search_sites("Engineering")` → site id
- `sharepoint_list_site_drives("<site-id>")` → drive id
- `sharepoint_list_children("<drive-id>", "/Protokolle")`,
  `sharepoint_get_item("<drive-id>", "01ABC…")`,
  `sharepoint_download_file("<drive-id>", "/Protokolle/notiz.txt")`

Teams shortcut: `teams_get_channel_files_folder("<teamId>", "<channelId>")`
→ `driveId` + `folderId`, then straight into the file tools.

## 7. Tools (10)

- `sharepoint_search_sites` – sites by name/keyword (pass `"root"` to
  `sharepoint_get_site` for the tenant root site)
- `sharepoint_get_site` – metadata of one site
- `sharepoint_list_site_drives` – document libraries (drives) of a site
- `sharepoint_get_item` – metadata of one file or folder on a drive
- `sharepoint_list_children` – folder contents, folders first (default `root`, up to 200;
  pages up to 5000 children before sorting folders-first and applying top,
  so very large folders may omit entries beyond that scan limit)
- `sharepoint_search_files` – name/keyword search on a drive
- `sharepoint_download_file` – without `localPath`: text decoded (truncated at
  20000 chars), binary as base64; above `maxBytes` (default 768 KB, max 2097152)
  rejected with `attachment-too-large`. **With `localPath` (absolute host path):
  streams the bytes to disk instead** – the result carries metadata + path only,
  so large files bypass the model context. Downloads are atomic (temp file +
  move) and refuse to overwrite without `overwrite: true`
- `sharepoint_create_folder` – renames on name conflict instead of failing
- `sharepoint_upload_file` – `localPath` (absolute host path, read from disk so
  bytes bypass the model context) or small inline `contentText`/`contentBase64`.
  Local files up to 4194304 bytes use simple upload, larger ones (up to
  104857600 bytes) a resumable Graph upload session (`createUploadSession`,
  same v1.0 endpoint – no endpoint switch needed). Existing names always get a
  renamed copy (`name 1.ext`), never overwritten – on all upload paths.
  Above the tool limit: SharePoint UI
- `sharepoint_move_item` – rename and/or reparent; move back to undo

Large-file transfers never pass through the model: the model only sends short
path strings (like `cp` arguments), the host process moves the bytes. Paths must
be absolute. Note: like `cp`, a path-based upload can exfiltrate any host file
the process can read – only point it at files you intend to share.

## 8. Error codes (returned as `isError` results with a `Next:` hint)

`site-not-found`, `drive-not-found`, `item-not-found`, `invalid-request`,
`attachment-too-large`, `channel-files-folder-not-found` (Teams bridge),
`auth-misconfigured`, `auth-failed`, `access-denied`, `throttled`,
`conflict`, `service-unavailable`, `graph-error`.
Details + stack traces go to the server log (stderr) only, never to the client.

## 9. Troubleshooting

- `[access-denied]` → check the Sites/Files consent/scopes from section 2
  (delegated `Sites.Read.All` + `Files.ReadWrite`, consent again after adding permissions)
- `[site-not-found]` → `sharepoint_search_sites` again; site names change on rename
- `[drive-not-found]` → `sharepoint_list_site_drives` for the site, or
  `teams_get_channel_files_folder` for a channel library
- `[auth-failed]` headless → `Graph__DelegatedFlow=DeviceCode`
- App-Only configured → switch to delegated; this host rejects
  `AuthMode: AppOnly` at startup with a hint
- Auth/account-type issues (personal accounts, token version, public
  client flows) → same fixes as in the
  [Outlook troubleshooting](outlook.md#9-troubleshooting)
