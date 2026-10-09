using System.Reflection;

namespace MicrosoftMcp.Host.Setup;

/// <summary>Global <c>--help</c> / <c>--version</c> handling for bare host invocations.
/// The unified binary is primarily an MCP stdio server, but a human running
/// <c>microsoft-mcp --help</c> must get usage text instead of a hanging server.</summary>
public static class HostHelp
{
    /// <summary>True when the args request global help or version output (no server start).</summary>
    public static bool ShouldShowHelp(string[] args)
    {
        return args.Any(a => a.Equals("--help", StringComparison.OrdinalIgnoreCase)
            || a.Equals("-h", StringComparison.Ordinal)
            || a.Equals("-?", StringComparison.Ordinal)
            || a.Equals("help", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>True when the args request the binary version (no server start).</summary>
    public static bool ShouldShowVersion(string[] args)
    {
        return args.Any(a => a.Equals("--version", StringComparison.OrdinalIgnoreCase)
            || a.Equals("-v", StringComparison.Ordinal));
    }

    public static string UsageText()
    {
        return "microsoft-mcp – local MCP servers for Microsoft 365 via Microsoft Graph.\n"
            + "\n"
            + "Usage:\n"
            + "  microsoft-mcp [--servers outlook,calendar]   Start the MCP server (default: all domains)\n"
            + "  microsoft-mcp setup [--servers ...] [--account work|personal] [--auth delegated|apponly]\n"
            + "                      [--client vscode|claude|opencode|codex|openclaw|hermes|generic]\n"
            + "                      [--binary PATH] [--headless] [--json]\n"
            + "  microsoft-mcp doctor [--servers ...] [--json]\n"
            + "  microsoft-mcp policy migrate [--path PATH] [--write] [--json]\n"
            + "  microsoft-mcp --help | --version\n"
            + "\n"
            + "Domains: outlook,onedrive,calendar,teams,sharepoint,todo,planner (or MCP_SERVERS env).\n"
            + "Setup details: docs/setup.en.md (same content as docs/setup.de.md).";
    }

    public static string VersionText()
    {
        string? informational = Assembly
            .GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;
        return $"microsoft-mcp {informational ?? "0.0.0"}";
    }
}
