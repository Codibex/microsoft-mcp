using System.Reflection;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using MicrosoftMcp.Common;

namespace MicrosoftMcp.Host.Setup;

/// <summary>Dispatch für <c>setup</c> / <c>doctor</c> (startet nie den MCP-Server).</summary>
public static class SetupCli
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    /// <summary>True wenn args mit setup/doctor (oder deren Hilfe) beginnen.</summary>
    public static bool IsSetupCommand(string[] args)
    {
        return args.Length > 0
            && (args[0].Equals("setup", StringComparison.OrdinalIgnoreCase)
                || args[0].Equals("doctor", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Führt setup/doctor aus, gibt den Exit-Code zurück.</summary>
    public static int Run(string[] args)
    {
        string command = args[0].ToLowerInvariant();
        Dictionary<string, string?> opts = ParseOpts(args[1..]);
        if (Has(opts, "help") || Has(opts, "h"))
        {
            PrintHelp(command);
            return 0;
        }

        bool json = Has(opts, "json");
        string? serversRaw = Opt(opts, "servers") ?? Environment.GetEnvironmentVariable("MCP_SERVERS");

        IReadOnlyList<string> servers;
        try
        {
            servers = ServerSelection.Parse(serversRaw);
        }
        catch (InvalidOperationException ex)
        {
            return Fail(json, ex.Message + " Next: use --servers outlook,onedrive,calendar,teams.");
        }

        return command switch
        {
            "setup" => RunSetup(servers, opts, json),
            "doctor" => RunDoctor(servers, json),
            _ => Fail(json, $"Unknown command '{command}'. Next: use setup or doctor."),
        };
    }

    private static int RunSetup(IReadOnlyList<string> servers, Dictionary<string, string?> opts, bool json)
    {
        string account = Opt(opts, "account") ?? "work";
        if (!account.Equals("work", StringComparison.OrdinalIgnoreCase)
            && !account.Equals("personal", StringComparison.OrdinalIgnoreCase))
        {
            return Fail(json, $"Unknown --account '{account}'. Next: use work or personal.");
        }

        string authRaw = Opt(opts, "auth") ?? "delegated";
        AuthMode auth = authRaw.ToLowerInvariant() switch
        {
            "delegated" => AuthMode.Delegated,
            "apponly" => AuthMode.AppOnly,
            "app-only" => AuthMode.AppOnly,
            _ => (AuthMode)(-1)
        };
        if ((int)auth < 0)
        {
            return Fail(json, $"Unknown --auth '{authRaw}'. Next: use delegated or apponly.");
        }

        string? comboError = SetupGuide.ValidateCombination(servers, auth);
        if (comboError is not null)
        {
            return Fail(json, comboError);
        }

        SetupClient client = SetupGuide.ParseClient(Opt(opts, "client"));
        string binary = Opt(opts, "binary") ?? "/path/to/microsoft-mcp";
        bool headless = Has(opts, "headless");
        IReadOnlyList<string> scopes = SetupGuide.ScopesFor(servers);
        string tenantHint = SetupGuide.TenantHint(account);
        string snippet = SetupGuide.BuildMcpJson(client, binary, servers, tenantHint, auth, headless);

        if (json)
        {
            Console.Out.WriteLine(JsonSerializer.Serialize(new
            {
                servers,
                scopes,
                account,
                auth = auth.ToString(),
                client = client.ToString(),
                clientKey = SetupGuide.TopLevelKey(client),
                tenantHint,
                mcpJson = snippet
            }, JsonOptions));
            return 0;
        }

        Console.Out.WriteLine($"Servers: {string.Join(",", servers)} | account: {account} | auth: {auth}");
        Console.Out.WriteLine($"TenantId: {(account.Equals("personal", StringComparison.OrdinalIgnoreCase) ? "common" : "<tenant-guid>")} | ClientId: <client-id from Entra overview>");
        Console.Out.WriteLine($"Delegated scopes to consent: {string.Join(" ", scopes)}"
            + (auth == AuthMode.AppOnly ? " (+ application Mail.ReadWrite, admin consent, client secret)" : ""));
        if (headless)
        {
            Console.Out.WriteLine("Headless: set Graph__DelegatedFlow=DeviceCode and confirm the code from stderr.");
        }

        Console.Out.WriteLine("Entra (once): app registration → Mobile+desktop http://localhost → API permissions above → consent.");
        Console.Out.WriteLine("Personal accounts: token v2 + AzureADandPersonalMicrosoftAccount (see outlook.md troubleshooting).");
        Console.Out.WriteLine();
        Console.Out.WriteLine(snippet);
        Console.Out.WriteLine();
        Console.Out.WriteLine("Verify: microsoft-mcp doctor --servers " + string.Join(",", servers));
        return 0;
    }

    private static int RunDoctor(IReadOnlyList<string> servers, bool json)
    {
        IConfiguration config = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: false)
            .AddEnvironmentVariables()
            .AddUserSecrets(Assembly.GetExecutingAssembly(), optional: true)
            .Build();

        GraphAuthOptions options = new();
        config.GetSection(GraphAuthOptions.SectionName).Bind(options);
        // Leere DelegatedScopes nicht auffüllen: DoctorChecks meldet dann
        // "Scopes follow the selection" mit der Union (wie der Host anfordert).

        string? policyPath = null;
        string? policyError = null;
        try
        {
            policyPath = PolicyFile.FindPolicyFile();
        }
        catch (Exception ex) when (ex is Microsoft.Extensions.Options.OptionsValidationException or UnauthorizedAccessException)
        {
            policyError = ex.Message;
        }

        IReadOnlyList<SetupCheck> checks = DoctorChecks.Run(servers, options, policyPath, policyError);

        if (json)
        {
            Console.Out.WriteLine(JsonSerializer.Serialize(new { servers, checks }, JsonOptions));
        }
        else
        {
            foreach (SetupCheck c in checks)
            {
                Console.Out.WriteLine($"{(c.Ok ? "ok  " : "FAIL")} {c.Id}: {c.Message}"
                    + (c.Next is null ? string.Empty : $" {c.Next}"));
            }
        }

        return checks.All(c => c.Ok) ? 0 : 1;
    }

    private static int Fail(bool json, string message)
    {
        if (json)
        {
            Console.Out.WriteLine(JsonSerializer.Serialize(new { error = message }, JsonOptions));
        }
        else
        {
            Console.Out.WriteLine(message);
        }

        return 2;
    }

    private static void PrintHelp(string command)
    {
        Console.Out.WriteLine(command == "setup"
            ? "Usage: microsoft-mcp setup [--servers outlook,calendar] [--account work|personal] [--auth delegated|apponly] [--client vscode|claude|generic] [--binary PATH] [--headless] [--json]"
            : "Usage: microsoft-mcp doctor [--servers outlook,calendar] [--json]");
    }

    private static Dictionary<string, string?> ParseOpts(string[] args)
    {
        Dictionary<string, string?> opts = new(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < args.Length; i++)
        {
            string a = args[i];
            if (a == "--")
            {
                continue;
            }

            if (a.StartsWith("--", StringComparison.Ordinal))
            {
                string body = a[2..];
                int eq = body.IndexOf('=');
                if (eq >= 0)
                {
                    opts[body[..eq]] = body[(eq + 1)..];
                }
                else if (i + 1 < args.Length && !args[i + 1].StartsWith('-'))
                {
                    opts[body] = args[i + 1];
                    i++;
                }
                else
                {
                    opts[body] = "true";
                }
            }
            else if (a.StartsWith('-') && a.Length == 2)
            {
                opts[a[1..]] = "true";
            }
        }

        return opts;
    }

    private static string? Opt(Dictionary<string, string?> opts, string name)
    {
        return opts.TryGetValue(name, out string? v) ? v : null;
    }

    private static bool Has(Dictionary<string, string?> opts, string name)
    {
        return opts.ContainsKey(name);
    }
}
