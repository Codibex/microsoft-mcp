using AwesomeAssertions;
using MicrosoftMcp.Common;

namespace MicrosoftMcp.Common.Tests;

public sealed class TokenCredentialFactoryTests
{
    private readonly TokenCredentialFactory _sut = new();

    [Fact]
    public void Null_options_throws()
    {
        Action act = () => _sut.GetCredential(null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void Delegated_without_tenant_throws()
    {
        var options = new GraphAuthOptions { AuthMode = AuthMode.Delegated, ClientId = "c" };
        Action act = () => _sut.GetCredential(options);
        act.Should().Throw<MailServiceException>().Where(e => e.Code == "auth-misconfigured").WithMessage("*TenantId*");
    }

    [Fact]
    public void Delegated_without_client_throws()
    {
        var options = new GraphAuthOptions { AuthMode = AuthMode.Delegated, TenantId = "t" };
        Action act = () => _sut.GetCredential(options);
        act.Should().Throw<MailServiceException>().Where(e => e.Code == "auth-misconfigured").WithMessage("*ClientId*");
    }

    [Fact]
    public void Delegated_devicecode_returns_credential_without_network()
    {
        var options = new GraphAuthOptions
        {
            AuthMode = AuthMode.Delegated,
            DelegatedFlow = DelegatedFlow.DeviceCode,
            TenantId = "t",
            ClientId = "c"
        };

        _sut.GetCredential(options).Should().NotBeNull();
    }

    [Fact]
    public void AppOnly_without_secret_throws()
    {
        var options = new GraphAuthOptions
        {
            AuthMode = AuthMode.AppOnly,
            TenantId = "t",
            ClientId = "c",
            UserIdOrUpn = "user@tenant"
        };

        Action act = () => _sut.GetCredential(options);
        act.Should().Throw<MailServiceException>().Where(e => e.Code == "auth-misconfigured").WithMessage("*ClientSecret*");
    }

    [Fact]
    public void AppOnly_certificate_is_not_supported_yet()
    {
        var options = new GraphAuthOptions
        {
            AuthMode = AuthMode.AppOnly,
            AppCredential = AppCredentialKind.Certificate,
            TenantId = "t",
            ClientId = "c",
            UserIdOrUpn = "user@tenant"
        };

        Action act = () => _sut.GetCredential(options);
        act.Should().Throw<NotSupportedException>();
    }

    [Fact]
    public void AppOnly_secret_returns_credential_without_network()
    {
        var options = new GraphAuthOptions
        {
            AuthMode = AuthMode.AppOnly,
            TenantId = "t",
            ClientId = "c",
            UserIdOrUpn = "user@tenant",
            ClientSecret = "s"
        };

        _sut.GetCredential(options).Should().NotBeNull();
    }
}
