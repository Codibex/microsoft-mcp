using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using MicrosoftMcp.Common;
using MicrosoftMcp.Outlook;
using ModelContextProtocol.Server;

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
    // (A "GRAPH_" prefix would strip to root-level keys and never bind.)
    .AddEnvironmentVariables()
    .AddUserSecrets<Program>();

builder.Services
    .AddGraphCommon(builder.Configuration)
    // Admin-owned messaging policy (policy.json) is the ONLY source for
    // recipient/disclosure rules — Messaging__* env vars are ignored by design.
    .AddOutlook(MessagingPolicySetup.Initialize());

builder.Services
    .AddMcpServer()
    .WithStdioServerTransport()
    .WithToolsFromAssembly(typeof(OutlookTools).Assembly);

var app = builder.Build();

// Fail fast on misconfiguration (validator runs here) and print the
// effective mode to stderr (always visible: the file log level is Warning).
var graph = app.Services.GetRequiredService<IOptions<GraphAuthOptions>>().Value;
Console.Error.WriteLine(
    $"[startup] Graph auth mode: {graph.AuthMode} (mailbox {graph.UserIdOrUpn}, flow {graph.DelegatedFlow})");

await app.RunAsync();
