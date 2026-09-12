using AwesomeAssertions;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;

namespace MicrosoftMcp.Common.Tests;

/// <summary>Guards the documented config contract: Graph__* env vars bind to
/// the "Graph" section (regression test: a provider prefix like "GRAPH_"
/// would strip to root-level keys and silently never bind).</summary>
public sealed class GraphConfigurationBindingTests
{
    [Fact]
    public void Double_underscore_env_vars_bind_to_graph_section()
    {
        string? backupTenant = Environment.GetEnvironmentVariable("Graph__TenantId");
        string? backupClient = Environment.GetEnvironmentVariable("Graph__ClientId");
        string? backupFlow = Environment.GetEnvironmentVariable("Graph__DelegatedFlow");
        try
        {
            Environment.SetEnvironmentVariable("Graph__TenantId", "t-env");
            Environment.SetEnvironmentVariable("Graph__ClientId", "c-env");
            Environment.SetEnvironmentVariable("Graph__DelegatedFlow", "DeviceCode");

            var config = new ConfigurationBuilder().AddEnvironmentVariables().Build();
            var services = new ServiceCollection();
            services.AddGraphCommon(config);
            var options = services.BuildServiceProvider()
                .GetRequiredService<IOptions<GraphAuthOptions>>().Value;

            options.TenantId.Should().Be("t-env");
            options.ClientId.Should().Be("c-env");
            options.DelegatedFlow.Should().Be(DelegatedFlow.DeviceCode);
        }
        finally
        {
            Environment.SetEnvironmentVariable("Graph__TenantId", backupTenant);
            Environment.SetEnvironmentVariable("Graph__ClientId", backupClient);
            Environment.SetEnvironmentVariable("Graph__DelegatedFlow", backupFlow);
        }
    }
}
