# AGENTS.md

Machine notes for this repo. Source of truth for the consumer setup:
`docs/setup.en.md` / `docs/setup.de.md` (same content,
then `docs/outlook.md §9–§10`, `teams.md §8–§9`).

## Supporting setup (consumer view)

- Always clarify the 3 questions from `docs/setup.en.md §0` first
  (account, domains, client: vscode, claude, opencode, codex, openclaw,
  hermes). Never guess, never invent secrets.
- `microsoft-mcp setup …` generates a checklist + client snippet
  (JSON, Codex: TOML, Hermes: YAML),
  `microsoft-mcp doctor [--json]` verifies offline. Both start
  no MCP server and need no secrets.
- Combinations: App-Only is only valid with `outlook`; `onedrive`,
  `calendar`, `teams` require Delegated. For personal accounts, use
  `TenantId=consumers` with a personal-only app, or `common` with a mixed
  org+personal app.
- Secrets only in `env`/user-secrets, never in `appsettings.json`, never commit.
  `policy.json` is the only policy source (`Messaging__*` is ignored).

## Build / verify

```bash
dotnet build MicrosoftMcp.slnx
dotnet test MicrosoftMcp.slnx
microsoft-mcp doctor --servers outlook,calendar
```

- Don't touch releases (no version bump, no tag, no `docs/releases/*`).
- Keep changes as one PR package, separated by docs / host / tests.
