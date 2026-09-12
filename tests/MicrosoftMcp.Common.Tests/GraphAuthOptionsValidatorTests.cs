using MicrosoftMcp.Common;

namespace MicrosoftMcp.Common.Tests;

public sealed class GraphAuthOptionsValidatorTests
{
    [Fact]
    public void Valid_delegated_options_succeed()
    {
        var result = new GraphAuthOptionsValidator().Validate(null, new GraphAuthOptions
        {
            AuthMode = AuthMode.Delegated,
            TenantId = "t",
            ClientId = "c"
        });

        Assert.True(result.Succeeded);
    }

    [Fact]
    public void Missing_tenant_fails()
    {
        var result = new GraphAuthOptionsValidator().Validate(null, new GraphAuthOptions
        {
            TenantId = string.Empty,
            ClientId = "c"
        });

        Assert.False(result.Succeeded);
    }

    [Fact]
    public void Missing_client_fails()
    {
        var result = new GraphAuthOptionsValidator().Validate(null, new GraphAuthOptions
        {
            TenantId = "t",
            ClientId = "  "
        });

        Assert.False(result.Succeeded);
    }

    [Fact]
    public void AppOnly_with_me_fails()
    {
        var result = new GraphAuthOptionsValidator().Validate(null, new GraphAuthOptions
        {
            AuthMode = AuthMode.AppOnly,
            TenantId = "t",
            ClientId = "c",
            UserIdOrUpn = "me"
        });

        Assert.False(result.Succeeded);
    }

    [Fact]
    public void Valid_apponly_options_succeed()
    {
        var result = new GraphAuthOptionsValidator().Validate(null, new GraphAuthOptions
        {
            AuthMode = AuthMode.AppOnly,
            TenantId = "t",
            ClientId = "c",
            UserIdOrUpn = "user@tenant"
        });

        Assert.True(result.Succeeded);
    }
}
