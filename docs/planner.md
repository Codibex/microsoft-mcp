# Planner MCP (`microsoft-mcp-planner`)

Local MCP server (Stdio) for **Microsoft Planner** via Microsoft Graph.
Same pattern as the [Outlook server](outlook.md): Stdio, `isError`
results with `[code]` hints, delegated auth via the shared `Common`
library. The server is strictly **read-only**; no write operations are
exposed.

The use-case this unlocks: after "what meetings do I have?" comes "who
owns what?" – tasks assigned to you, plans of a team group, and the
details (description, checklist, assignees) behind a task.

## 1. Prerequisites

- .NET 10 SDK (`dotnet --version` → 10.x)
- Microsoft 365 Planner access (work/school account)

## 2. Entra setup (one time)

In the **same app registration** (or a new one), add the delegated
permissions (no admin needed in most tenants, otherwise ask yours):

**API permissions → Add → Microsoft Graph → Delegated:**
`Tasks.Read`, `Group.Read.All` (`offline_access` is added automatically).

Plans live on Microsoft 365 groups, so `Group.Read.All` is needed to
resolve `planner_list_plans(groupId)`. `Tasks.Read` covers the task
reads.

No redirect URI is needed for device-code flow; for browser login add
the **Mobile and desktop applications** platform with `http://localhost`
(see [Outlook guide](outlook.md), recipe A steps 1–3).

App-Only is **not** supported by this host: Planner runs via `/me/*`
and `/groups/*` and needs a signed-in user (service access would
require a different permission/path scheme – out of scope).

Work or school accounts only: group-backed Planner APIs are not
supported for personal Microsoft accounts, so `TenantId=consumers`
cannot work here – use your org tenant GUID.

## 3. Config (same 2 values as Outlook)

```bash
dotnet user-secrets set "Graph:TenantId" "<tenant-id>" \
  --project src/MicrosoftMcp.Planner.Host
dotnet user-secrets set "Graph:ClientId" "<client-id>" \
  --project src/MicrosoftMcp.Planner.Host
```

The host template defaults `DelegatedScopes` to the Planner scopes, so
no scope config is needed. Headless setups use
`Graph:DelegatedFlow` = `DeviceCode`.
Planner has no `policy.json` section because the server is read-only.

## 4. Run

```bash
dotnet run --project src/MicrosoftMcp.Planner.Host
```

Startup logs the mode (`[startup] Planner auth mode: Delegated …`) and
fails fast on missing values with a `Next:` hint.

## 5. Wire up an MCP client

```json
{
  "servers": {
    "planner": {
      "command": "/path/to/microsoft-mcp-planner",
      "env": {
        "Graph__TenantId": "<tenant-id>",
        "Graph__ClientId": "<client-id>"
      }
    }
  }
}
```

For production use take the published binary (or a release asset).

## 6. Tools (4)

- `planner_list_my_tasks` – tasks assigned to me
- `planner_list_plans` – plans of a Microsoft 365 group (`groupId`)
- `planner_list_plan_tasks` – tasks of a plan (`planId`)
- `planner_read_task` – one task with description, checklist,
  assignees and bucket name

Example prompt: "What am I owning this week
(`planner_list_my_tasks`)? Show the plan tasks
(`planner_list_plan_tasks`) and read the blocked one
(`planner_read_task`)."

## 7. Error codes (returned as `isError` results with a `Next:` hint)

`planner-plan-not-found`, `planner-task-not-found`, `group-not-found`,
`invalid-request`, `auth-misconfigured`, `auth-failed`,
`access-denied`, `throttled`, `conflict`, `service-unavailable`,
`graph-error`.
Details + stack traces go to the server log (stderr) only, never to the client.

## 8. Troubleshooting

- `[access-denied]` → check the Tasks/Group consent/scopes from
  section 2 (delegated `Tasks.Read` + `Group.Read.All`, consent again
  after adding permissions)
- `[group-not-found]` → verify the Microsoft 365 group id backing the
  plan (not a Teams team id or channel id)
- `[planner-plan-not-found]` → `planner_list_plans` again with the group id
- `[planner-task-not-found]` → `planner_list_my_tasks` or
  `planner_list_plan_tasks` again for a valid task id
- `[auth-failed]` headless → `Graph__DelegatedFlow=DeviceCode`
- App-Only configured → switch to delegated; this host rejects
  `AuthMode: AppOnly` at startup with a hint
- Auth/account-type issues (token version, public client flows) →
  same fixes as in the
  [Outlook troubleshooting](outlook.md#9-troubleshooting)
