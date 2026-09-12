# MicrosoftMcp

Local MCP servers (Stdio) for Microsoft 365 via Microsoft Graph.
.NET 10 / C# 14, central package management (`Directory.Packages.props`).

| Server | Binary | Guide | Status |
|---|---|---|---|
| Outlook email triage (17 tools, drafts only, no send) | `microsoft-mcp-outlook` | [docs/outlook.md](docs/outlook.md) | ✅ Live-verified (personal account) |
| OneDrive files (9 tools, read/create/move, no delete) | `microsoft-mcp-onedrive` | [docs/onedrive.md](docs/onedrive.md) | Implemented |

Shared foundations (delegated + app-only auth, `isError` results with
`[code]` + `Next:` hints, censoring logs) live in `src/MicrosoftMcp.Common`.

## Quickstart

```bash
dotnet build MicrosoftMcp.slnx
dotnet test MicrosoftMcp.slnx
```

Configuration via user-secrets or `Graph__*` env vars (never commit
secrets) – details per server in its guide. Each server logs its
effective auth mode on startup and fails fast on missing values.

## Layout

```text
src/MicrosoftMcp.Common/        # auth, Graph client, error mapping
src/MicrosoftMcp.Outlook/       # mail service + tools (lib)
src/MicrosoftMcp.Outlook.Host/  # microsoft-mcp-outlook (exe)
src/MicrosoftMcp.OneDrive/      # drive service + tools (lib)
src/MicrosoftMcp.OneDrive.Host/ # microsoft-mcp-onedrive (exe)
tests/                          # xUnit + NSubstitute, incl. in-memory fakes
```

## Versioning (SemVer)

`MAJOR.MINOR.PATCH`, centralized in `Directory.Build.props`
(`VersionPrefix`). While `MAJOR = 0`, `MINOR` may contain breaking changes:
new tool/feature → `MINOR`, fix → `PATCH`, breaking from `1.0.0` → `MAJOR`.
Push tag `vX.Y.Z` → GitHub Action builds, tests and releases binaries
(linux-x64, win-x64, osx-arm64) with checksums; `-rc.1` tags become prereleases.

## License

MIT – see [LICENSE](LICENSE).
