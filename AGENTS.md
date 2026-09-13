# AGENTS.md

Maschinen-Hinweise für dieses Repo. Fachliche Wahrheit für das
Verwender-Setup: `docs/setup.md` (danach `docs/outlook.md §9–§10`,
`teams.md §8–§9`).

## Setup unterstützen (Verwender-Sicht)

- Immer zuerst die 3 Fragen aus `docs/setup.md §0` klären
  (Konto, Domains, Client). Nie raten, nie Secrets erfinden.
- `microsoft-mcp setup …` generiert Checklist + `mcp.json`-Snippet,
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
