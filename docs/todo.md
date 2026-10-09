# To Do MCP (`microsoft-mcp-todo`)

Local MCP server (Stdio) for **Microsoft To Do** via Microsoft Graph.
Same pattern as the [Outlook server](outlook.md): Stdio, `isError`
results with `[code]` hints, delegated auth via the shared `Common`
library. Add + complete only – **no delete** (same spirit as the mail
drafts-only rule).

The use-case this unlocks: an agent that reads mail and meetings can
*harvest commitments* ("I'll send you the offer Friday") and file them
as tracked To Do items the user sees in their own To Do app – visible,
editable, syncs to the phone.

## 1. Prerequisites

- .NET 10 SDK (`dotnet --version` → 10.x)
- A Microsoft 365 / Microsoft account with To Do

## 2. Entra setup (one time)

In the **same app registration** (or a new one), add the delegated
permissions (no admin needed in most tenants, otherwise ask yours):

**API permissions → Add → Microsoft Graph → Delegated:**
`Tasks.Read`, `Tasks.ReadWrite`, `Tasks.Read.Shared`,
`Tasks.ReadWrite.Shared` (`offline_access` is added automatically).

The `.Shared` variants cover To Do lists shared with you (family/team
lists several people tick through) – without them, `/me/todo/lists`
only shows your own lists.

No redirect URI is needed for device-code flow; for browser login add
the **Mobile and desktop applications** platform with `http://localhost`
(see [Outlook guide](outlook.md), recipe A steps 1–3).

App-Only is **not** supported by this host: To Do runs via `/me/*` and
needs a signed-in user (service access would require a different
permission/path scheme – out of scope).

For supported personal-account scenarios, use `Graph:TenantId=consumers`
with a personal-only app, or `common` with an org + personal app (see
[Outlook troubleshooting](outlook.md#9-troubleshooting)). Headless:
`Graph:DelegatedFlow` = `DeviceCode`.

## 3. Config (same 2 values as Outlook)

```bash
dotnet user-secrets set "Graph:TenantId" "<tenant-id>" \
  --project src/MicrosoftMcp.Todo.Host
dotnet user-secrets set "Graph:ClientId" "<client-id>" \
  --project src/MicrosoftMcp.Todo.Host
```

The host template defaults `DelegatedScopes` to the To Do scopes, so no
scope config is needed. To Do has no `policy.json` section because its
tools do not send to recipients or invite attendees.

## 4. Run

```bash
dotnet run --project src/MicrosoftMcp.Todo.Host
```

Startup logs the mode (`[startup] To Do auth mode: Delegated …`) and
fails fast on missing values with a `Next:` hint.

## 5. Wire up an MCP client

```json
{
  "servers": {
    "todo": {
      "command": "/path/to/microsoft-mcp-todo",
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

- `todo_list_lists` – To Do lists with id, name and shared flag
- `todo_list_tasks` – tasks of a list (default: well-known Tasks list;
  completed tasks hidden unless `includeCompleted: true`)
- `todo_add_task` – add a task (`subject`, optional `listId`,
  `dueDateTime` ISO, `importance` low|normal|high, `body` notes)
- `todo_complete_task` – mark a task completed (`taskId`, optional
  `listId` to skip discovery)

Example prompt: "What did I promise this week? Check
`todo_list_tasks`, and file 'Send the offer by Friday' via
`todo_add_task` with a due date. Tick off what is done with
`todo_complete_task`."

## 7. Error codes (returned as `isError` results with a `Next:` hint)

`todo-list-not-found`, `todo-task-not-found`, `invalid-request`,
`auth-misconfigured`, `auth-failed`, `access-denied`, `throttled`,
`conflict`, `service-unavailable`, `graph-error`.
Details + stack traces go to the server log (stderr) only, never to the client.

## 8. Troubleshooting

- `[access-denied]` → check the Tasks consent/scopes from section 2
  (all four delegated scopes, consent again after adding permissions)
- `[todo-list-not-found]` → `todo_list_lists` again; list names change on rename
- `[todo-task-not-found]` → `todo_list_tasks` again (also with
  `includeCompleted: true` – completed tasks are hidden by default)
- `[auth-failed]` headless → `Graph__DelegatedFlow=DeviceCode`
- App-Only configured → switch to delegated; this host rejects
  `AuthMode: AppOnly` at startup with a hint
- Auth/account-type issues (personal accounts, token version, public
  client flows) → same fixes as in the
  [Outlook troubleshooting](outlook.md#9-troubleshooting)
