# AGENTS.md

Maschinen-Hinweise für dieses Repo. Fachliche Wahrheit für das
Verwender-Setup: `docs/setup.de.md` / `docs/setup.en.md` (inhaltsgleich,
danach `docs/outlook.md §9–§10`, `teams.md §8–§9`).

## Setup unterstützen (Verwender-Sicht)

- Immer zuerst die 3 Fragen aus `docs/setup.de.md §0` klären
  (Konto, Domains, Client: vscode, claude, opencode, codex, openclaw,
  hermes). Nie raten, nie Secrets erfinden.
- `microsoft-mcp setup …` generiert Checklist + Client-Snippet
  (JSON, Codex: TOML, Hermes: YAML),
  `microsoft-mcp doctor [--json]` verifiziert offline. Beide starten
  keinen MCP-Server und brauchen keine Secrets.
- Kombinationen: App-Only nur mit `outlook` gültig; `onedrive`,
  `calendar`, `teams` erfordern Delegated. Personal → `TenantId=common`.
- Secrets nur in `env`/User-Secrets, nie in `appsettings.json`, nie committen.
  `policy.json` ist die einzige Policy-Quelle (`Messaging__*` wird ignoriert).

## Bauen / Prüfen

```bash
dotnet build MicrosoftMcp.slnx
dotnet test MicrosoftMcp.slnx
microsoft-mcp doctor --servers outlook,calendar
```

- Keine Releases anfassen (kein Version-Bump, kein Tag, keine `docs/releases/*`).
- Änderungen als ein PR-Paket halten, getrennt nach Doku / Host / Tests.
