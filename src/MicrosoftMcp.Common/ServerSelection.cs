namespace MicrosoftMcp.Common;

/// <summary>Selects which domain toolsets a unified host process exposes.
/// Keeps setup to one binary/config while preserving per-domain isolation
/// (run separate processes with different selections) and least-privilege
/// scopes (requested scopes follow the selection).</summary>
public static class ServerSelection
{
    public static readonly IReadOnlyList<string> All = ["outlook", "onedrive", "calendar", "teams"];

    private static readonly IReadOnlyDictionary<string, string[]> DefaultScopes =
        new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
        {
            ["outlook"] = ["Mail.Read", "Mail.ReadWrite"],
            ["onedrive"] = ["Files.Read", "Files.ReadWrite"],
            ["calendar"] = ["Calendars.Read"],
            ["teams"] = [
                "Team.ReadBasic.All", "ChannelMessage.Read.All", "Chat.ReadBasic", "Chat.Read",
                "OnlineMeetingTranscript.Read.All", "OnlineMeetingAiInsight.Read.All"
            ]
        };

    /// <summary>Parses "--servers outlook,calendar" / MCP_SERVERS. Null or
    /// empty selects all (convenience default; an explicit selection narrows
    /// tools and scopes). Unknown names throw with the valid list.</summary>
    public static IReadOnlyList<string> Parse(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return All;
        }

        var selected = raw.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var unknown = selected.Where(s => !All.Contains(s, StringComparer.OrdinalIgnoreCase)).ToList();
        if (unknown.Count > 0)
        {
            throw new InvalidOperationException(
                $"Unknown server(s): {string.Join(", ", unknown)}. Valid: {string.Join(", ", All)}.");
        }

        return [.. All.Where(a => selected.Contains(a, StringComparer.OrdinalIgnoreCase))];
    }

    /// <summary>Union of default scopes for the selection, in stable order.</summary>
    public static IReadOnlyList<string> DefaultScopesFor(IEnumerable<string> servers)
    {
        var scopes = new List<string>();
        foreach (string server in All)
        {
            if (servers.Contains(server, StringComparer.OrdinalIgnoreCase)
                && DefaultScopes.TryGetValue(server, out var own))
            {
                scopes.AddRange(own.Where(s => !scopes.Contains(s)));
            }
        }

        return scopes;
    }
}
