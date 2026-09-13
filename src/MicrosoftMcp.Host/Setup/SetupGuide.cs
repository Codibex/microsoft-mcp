using MicrosoftMcp.Common;

namespace MicrosoftMcp.Host.Setup;

/// <summary>Client-Typen für das generierte <c>mcp.json</c>-Snippet.</summary>
public enum SetupClient
{
    Vscode,
    Claude,
    Generic
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
            "vscode" => SetupClient.Vscode,
            "claude" => SetupClient.Claude,
            _ => SetupClient.Generic,
        };
    }

    /// <summary>Top-Level-Schlüssel je Client (vscode/generic: servers, claude: mcpServers).</summary>
    public static string TopLevelKey(SetupClient client)
    {
        return client == SetupClient.Claude ? "mcpServers" : "servers";
    }

    /// <summary>Baut das kopierfertige mcp.json-Snippet (keine Secrets enthalten).</summary>
    public static string BuildMcpJson(
        SetupClient client,
        string binaryPath,
        IReadOnlyList<string> servers,
        string tenantPlaceholder,
        AuthMode auth,
        bool headless)
    {
        string key = TopLevelKey(client);
        string serversArg = string.Join(",", servers);
        List<string> envLines =
        [
            $"          \"Graph__TenantId\": \"{tenantPlaceholder}\",",
            "          \"Graph__ClientId\": \"<client-id>\""
        ];
        if (auth == AuthMode.AppOnly)
        {
            envLines.Add("          ,\"Graph__AuthMode\": \"AppOnly\",");
            envLines.Add("          \"Graph__UserIdOrUpn\": \"mailbox@example.com\",");
            envLines.Add("          \"Graph__ClientSecret\": \"<secret>\"");
        }

        if (headless)
        {
            envLines.Add("          ,\"Graph__DelegatedFlow\": \"DeviceCode\"");
        }

        string env = string.Join("\n", envLines).Replace(",\n          ,", ",\n          ");
        return "{\n"
            + $"  \"{key}\": {{\n"
            + "    \"m365\": {\n"
            + $"      \"command\": \"{binaryPath}\",\n"
            + $"      \"args\": [\"--servers\", \"{serversArg}\"],\n"
            + "      \"env\": {\n"
            + env + "\n"
            + "      }\n"
            + "    }\n"
            + "  }\n"
            + "}";
    }
}
