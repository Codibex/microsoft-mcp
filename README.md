# MicrosoftMcp

Local MCP servers (Stdio) for Microsoft 365 via Microsoft Graph.
.NET 10 / C# 14, central package management (`Directory.Packages.props`).

| Server | Binary | Guide | Status |
|---|---|---|---|
| Outlook email triage (17 tools, drafts only, no send) | `microsoft-mcp-outlook` | [docs/outlook.md](docs/outlook.md) | ✅ Live-verified (personal account) |
| OneDrive files (9 tools, read/create/move, no delete) | `microsoft-mcp-onedrive` | [docs/onedrive.md](docs/onedrive.md) | Implemented |
| Calendar (6 tools, delegated writes) | `microsoft-mcp-calendar` | [docs/calendar.md](docs/calendar.md) | Implemented |
| Teams (14 tools, guarded message sends) | `microsoft-mcp-teams` | [docs/teams.md](docs/teams.md) | Implemented |

Shared foundations (delegated + app-only auth, `isError` results with
`[code]` + `Next:` hints, censoring logs) live in `src/MicrosoftMcp.Common`.

## Unified host (alternative)

One binary `microsoft-mcp` (`src/MicrosoftMcp.Host`) with all 46 tools,
domains selectable per process: `microsoft-mcp --servers outlook,calendar`
(or `MCP_SERVERS` env; default without arguments: all). Requested scopes
follow the selection unless `Graph:DelegatedScopes` is set explicitly.
Same Entra app and secrets work; user-secrets id is `microsoft-mcp-dev`.

```json
{
  "servers": {
    "m365": {
      "command": "/path/to/microsoft-mcp",
      "env": { "Graph__TenantId": "<id>", "Graph__ClientId": "<id>" }
    },
    "outlook-only": {
      "command": "/path/to/microsoft-mcp",
      "args": ["--servers", "outlook"],
      "env": { "Graph__TenantId": "<id>", "Graph__ClientId": "<id>" }
    }
  }
}
```

## Quickstart

```bash
dotnet build MicrosoftMcp.slnx
dotnet test MicrosoftMcp.slnx
```

Configuration via user-secrets or `Graph__*` env vars (never commit
secrets) – start with `docs/setup.de.md` / `docs/setup.en.md` (consumer setup:
binary, Entra app, client config for VS Code, Claude, OpenCode, Codex,
OpenClaw, Hermes, `doctor`), details per server in its guide. Each server logs its
effective auth mode on startup and fails fast on missing values.

## Layout

```text
src/MicrosoftMcp.Common/        # auth, Graph client, error mapping
src/MicrosoftMcp.Host/          # microsoft-mcp (exe, --servers selection)
src/MicrosoftMcp.Outlook/       # mail service + tools (lib)
src/MicrosoftMcp.Outlook.Host/  # microsoft-mcp-outlook (exe)
src/MicrosoftMcp.OneDrive/      # drive service + tools (lib)
src/MicrosoftMcp.OneDrive.Host/ # microsoft-mcp-onedrive (exe)
src/MicrosoftMcp.Calendar/      # calendar service + tools (lib)
src/MicrosoftMcp.Calendar.Host/ # microsoft-mcp-calendar (exe)
src/MicrosoftMcp.Teams/         # teams service + tools (lib)
src/MicrosoftMcp.Teams.Host/    # microsoft-mcp-teams (exe)
tests/                          # xUnit + NSubstitute, incl. in-memory fakes
```

## Versioning (SemVer)

`MAJOR.MINOR.PATCH`, centralized in `Directory.Build.props`
(`VersionPrefix`). While `MAJOR = 0`, `MINOR` may contain breaking changes:
new tool/feature → `MINOR`, fix → `PATCH`, breaking from `1.0.0` → `MAJOR`.
Push tag `vX.Y.Z` → GitHub Action builds, tests and releases binaries
(linux-x64/arm64, win-x64/arm64, osx-arm64) with checksums; `-rc.1` tags become prereleases.
Release notes are curated in `docs/releases/<tag>.md` and used as the
GitHub release body (auto-generated commit lists are appended).

## Protocol

MCP C# SDK 2.2.0. The server negotiates up to **2025-11-25** via the
classic `initialize` handshake and advertises **2026-07-28** via the new
`server/discover` flow with per-request `_meta` metadata. Tool names are
globally unique with domain prefixes (`outlook_*`, `onedrive_*`,
`calendar_*`, `teams_*`); the spec requires uniqueness within a server
and recommends identifier-prefixing for aggregated domains.

## License

MIT – see [LICENSE](LICENSE).
