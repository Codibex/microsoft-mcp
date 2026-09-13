using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using MicrosoftMcp.Calendar;
using MicrosoftMcp.Common;
using MicrosoftMcp.OneDrive;
using MicrosoftMcp.Outlook;
using MicrosoftMcp.Teams;
using ModelContextProtocol.Server;

// One binary, selectable domains: --servers outlook,calendar (or MCP_SERVERS env).
// Run separate processes per domain for isolation, or one process for everything.
string? serversArg = ArgValue(args, "--servers")
    ?? Environment.GetEnvironmentVariable("MCP_SERVERS");

IReadOnlyList<string> enabled;
try
{
    enabled = ServerSelection.Parse(serversArg);
}
catch (InvalidOperationException ex)
{
    Console.Error.WriteLine($"[startup] {ex.Message}");
    return 1;
}

var builder = Host.CreateApplicationBuilder(new HostApplicationBuilderSettings
{
    Args = args,
    // MCP clients launch the server with arbitrary cwd: always resolve
    // appsettings.json / user-secrets relative to the binary location.
    ContentRootPath = AppContext.BaseDirectory
});

// MCP speaks JSON-RPC on stdout -> all logs must go to stderr.
builder.Logging.AddConsole(o => o.LogToStandardErrorThreshold = LogLevel.Trace);

builder.Configuration
    .AddJsonFile("appsettings.json", optional: true, reloadOnChange: false)
    // No prefix: Graph__TenantId binds to the "Graph" section ("Graph:TenantId").
    .AddEnvironmentVariables()
    .AddUserSecrets<Program>();

builder.Services.AddGraphCommon(builder.Configuration);

// Requested scopes follow the selection (least privilege), unless the user
// set Graph:DelegatedScopes explicitly (then that wins).
IReadOnlyList<string> union = ServerSelection.DefaultScopesFor(enabled);
builder.Services.Configure<GraphAuthOptions>(o =>
{
    if (o.DelegatedScopes.Length == 0)
    {
        o.DelegatedScopes = [.. union];
    }
});

var mcp = builder.Services.AddMcpServer().WithStdioServerTransport();

// Shared admin-owned messaging policy (policy.json only, never env).
// Loaded once for all messaging domains; validated + audit-logged here.
MessagingPolicyOptions? messagingPolicy = null;
if (enabled.Contains("outlook") || enabled.Contains("teams"))
{
    messagingPolicy = MessagingPolicySetup.Initialize();
}

if (enabled.Contains("outlook"))
{
    builder.Services.AddOutlook(messagingPolicy);
    mcp.WithToolsFromAssembly(typeof(OutlookTools).Assembly);
}

if (enabled.Contains("onedrive"))
{
    builder.Services.AddOneDrive();
    mcp.WithToolsFromAssembly(typeof(DriveTools).Assembly);
}

if (enabled.Contains("calendar"))
{
    builder.Services.AddCalendar();
    mcp.WithToolsFromAssembly(typeof(CalendarTools).Assembly);
}

if (enabled.Contains("teams"))
{
    builder.Services.AddTeams();
    mcp.WithToolsFromAssembly(typeof(TeamsTools).Assembly);
}

var app = builder.Build();

// Fail fast on misconfiguration (validator runs here) and print the
// effective setup to stderr (always visible: the file log level is Warning).
var graph = app.Services.GetRequiredService<IOptions<GraphAuthOptions>>().Value;
Console.Error.WriteLine(
    $"[startup] Servers: {string.Join(",", enabled)} | auth mode: {graph.AuthMode} | scopes: {string.Join(" ", graph.DelegatedScopes)}");

await app.RunAsync();
return 0;

static string? ArgValue(string[] args, string name)
{
    for (int i = 0; i < args.Length; i++)
    {
        if (args[i] == name && i + 1 < args.Length)
        {
            return args[i + 1];
        }

        if (args[i].StartsWith(name + "=", StringComparison.Ordinal))
        {
            return args[i][(name.Length + 1)..];
        }
    }

    return null;
}
