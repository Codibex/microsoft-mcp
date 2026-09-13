using System.Text;
using MicrosoftMcp.Common;

namespace MicrosoftMcp.Host.Setup;

/// <summary>Client-Typen für das generierte Config-Snippet.</summary>
public enum SetupClient
{
    Vscode,
    Claude,
    Opencode,
    Codex,
    Openclaw,
    Hermes,
    Generic
}

/// <summary>Serialisierungsformat des Snippets.</summary>
public enum SnippetFormat
{
    Json,
    Toml,
    Yaml
}

/// <summary>Reine Setup-Logik (ohne I/O): Kombination prüfen, Scopes auflösen, Snippet bauen.</summary>
public static class SetupGuide
{
    /// <summary>Prüft AuthMode gegen die gewählten Domains. Null = gültig.</summary>
    public static string? ValidateCombination(IReadOnlyList<string> servers, AuthMode auth)
    {
        if (auth == AuthMode.AppOnly
            && servers.Any(s => !s.Equals("outlook", StringComparison.OrdinalIgnoreCase)))
        {
            return "AppOnly is only valid with outlook. Next: use --auth delegated for onedrive, calendar, teams (/me/* APIs require a signed-in user).";
        }

        return null;
    }

    /// <summary>Scope-Union für die Auswahl (Least Privilege).</summary>
    public static IReadOnlyList<string> ScopesFor(IReadOnlyList<string> servers)
    {
        return ServerSelection.DefaultScopesFor(servers);
    }

    /// <summary>Tenant-Platzhalter je Kontotyp.</summary>
    public static string TenantHint(string account)
    {
        return account.Equals("personal", StringComparison.OrdinalIgnoreCase)
            ? "common"
            : "<tenant-guid>";
    }

    /// <summary>Parst den Client-Namen, unbekannt fällt auf generic zurück.</summary>
    public static SetupClient ParseClient(string? raw)
    {
        return raw?.ToLowerInvariant() switch
        {
            "vscode" or "vs-code" or "code" => SetupClient.Vscode,
            "claude" or "claude-desktop" => SetupClient.Claude,
            "opencode" or "open-code" => SetupClient.Opencode,
            "codex" or "codex-cli" => SetupClient.Codex,
            "openclaw" or "open-claw" or "claw" => SetupClient.Openclaw,
            "hermes" or "hermes-agent" => SetupClient.Hermes,
            _ => SetupClient.Generic,
        };
    }

    /// <summary>Top-Level-Schlüssel je Client (z. B. vscode: servers, claude: mcpServers).</summary>
    public static string TopLevelKey(SetupClient client)
    {
        return client switch
        {
            SetupClient.Claude => "mcpServers",
            SetupClient.Opencode => "mcp",
            SetupClient.Codex => "mcp_servers",
            SetupClient.Openclaw => "mcp.servers",
            SetupClient.Hermes => "mcp_servers",
            _ => "servers",
        };
    }

    /// <summary>Config-Datei je Client.</summary>
    public static string ConfigFile(SetupClient client)
    {
        return client switch
        {
            SetupClient.Vscode => ".vscode/mcp.json",
            SetupClient.Claude => "claude_desktop_config.json",
            SetupClient.Opencode => "opencode.json",
            SetupClient.Codex => "~/.codex/config.toml",
            SetupClient.Openclaw => "openclaw.json (mcp.servers)",
            SetupClient.Hermes => "~/.hermes/config.yaml",
            _ => "mcp.json",
        };
    }

    /// <summary>Serialisierungsformat je Client (codex: TOML, hermes: YAML, Rest: JSON).</summary>
    public static SnippetFormat Format(SetupClient client)
    {
        return client switch
        {
            SetupClient.Codex => SnippetFormat.Toml,
            SetupClient.Hermes => SnippetFormat.Yaml,
            _ => SnippetFormat.Json,
        };
    }

    /// <summary>Env-Einträge (Name + Platzhalter) für die Auswahl.</summary>
    public static IReadOnlyList<(string Name, string Value)> EnvEntries(
        string tenantPlaceholder, AuthMode auth, bool headless)
    {
        List<(string Name, string Value)> entries =
        [
            ("Graph__TenantId", tenantPlaceholder),
            ("Graph__ClientId", "<client-id>")
        ];
        if (auth == AuthMode.AppOnly)
        {
            entries.Add(("Graph__AuthMode", "AppOnly"));
            entries.Add(("Graph__UserIdOrUpn", "mailbox@example.com"));
            entries.Add(("Graph__ClientSecret", "<secret>"));
        }

        if (headless)
        {
            entries.Add(("Graph__DelegatedFlow", "DeviceCode"));
        }

        return entries;
    }

    /// <summary>Baut das kopierfertige Config-Snippet (keine Secrets enthalten).</summary>
    public static string BuildMcpJson(
        SetupClient client,
        string binaryPath,
        IReadOnlyList<string> servers,
        string tenantPlaceholder,
        AuthMode auth,
        bool headless)
    {
        IReadOnlyList<(string Name, string Value)> env = EnvEntries(tenantPlaceholder, auth, headless);
        string serversArg = string.Join(",", servers);
        return Format(client) switch
        {
            SnippetFormat.Toml => BuildToml(binaryPath, serversArg, env),
            SnippetFormat.Yaml => BuildYaml(binaryPath, serversArg, env),
            _ => BuildJson(client, binaryPath, serversArg, env),
        };
    }

    private static string BuildJson(
        SetupClient client, string binaryPath, string serversArg,
        IReadOnlyList<(string Name, string Value)> env)
    {
        string args = $"[\"--servers\", \"{serversArg}\"]";
        string EnvBlock(string indent)
        {
            return string.Join(",\n", env.Select(e => $"{indent}\"{e.Name}\": \"{e.Value}\""));
        }

        // OpenCode: command ist ein Array, Env heißt "environment".
        if (client == SetupClient.Opencode)
        {
            return "{\n"
                + "  \"mcp\": {\n"
                + "    \"m365\": {\n"
                + "      \"type\": \"local\",\n"
                + $"      \"command\": [\"{binaryPath}\", \"--servers\", \"{serversArg}\"],\n"
                + "      \"environment\": {\n"
                + EnvBlock("        ") + "\n"
                + "      }\n"
                + "    }\n"
                + "  }\n"
                + "}";
        }

        // OpenClaw: mcp.servers mit enabled-Flag.
        if (client == SetupClient.Openclaw)
        {
            return "{\n"
                + "  \"mcp\": {\n"
                + "    \"servers\": {\n"
                + "      \"m365\": {\n"
                + $"        \"command\": \"{binaryPath}\",\n"
                + $"        \"args\": {args},\n"
                + "        \"env\": {\n"
                + EnvBlock("          ") + "\n"
                + "        },\n"
                + "        \"enabled\": true\n"
                + "      }\n"
                + "    }\n"
                + "  }\n"
                + "}";
        }

        string key = TopLevelKey(client);
        return "{\n"
            + $"  \"{key}\": {{\n"
            + "    \"m365\": {\n"
            + $"      \"command\": \"{binaryPath}\",\n"
            + $"      \"args\": {args},\n"
            + "      \"env\": {\n"
            + EnvBlock("        ") + "\n"
            + "      }\n"
            + "    }\n"
            + "  }\n"
            + "}";
    }

    private static string BuildToml(
        string binaryPath, string serversArg,
        IReadOnlyList<(string Name, string Value)> env)
    {
        StringBuilder sb = new();
        sb.AppendLine("[mcp_servers.m365]");
        sb.AppendLine($"command = \"{binaryPath}\"");
        sb.AppendLine($"args = [\"--servers\", \"{serversArg}\"]");
        sb.AppendLine();
        sb.AppendLine("[mcp_servers.m365.env]");
        foreach ((string name, string value) in env)
        {
            sb.AppendLine($"{name} = \"{value}\"");
        }

        return sb.ToString().TrimEnd();
    }

    private static string BuildYaml(
        string binaryPath, string serversArg,
        IReadOnlyList<(string Name, string Value)> env)
    {
        StringBuilder sb = new();
        sb.AppendLine("mcp_servers:");
        sb.AppendLine("  m365:");
        sb.AppendLine($"    command: \"{binaryPath}\"");
        sb.AppendLine($"    args: [\"--servers\", \"{serversArg}\"]");
        sb.AppendLine("    env:");
        foreach ((string name, string value) in env)
        {
            sb.AppendLine($"      {name}: \"{value}\"");
        }

        return sb.ToString().TrimEnd();
    }
}
