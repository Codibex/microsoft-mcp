using Azure.Core;
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

    [Fact]
    public void Token_cache_fallback_switches_to_memory_and_warns_once()
    {
        List<string> warnings = [];
        var persistent = new FailingCredential("Persistence check failed: libsecret");
        var memory = new CountingCredential();
        var credential = new TokenCacheCredential(persistent, memory, warnings.Add);
        var context = new TokenRequestContext(["scope"]);

        credential.GetToken(context, CancellationToken.None);
        credential.GetToken(context, CancellationToken.None);

        memory.Calls.Should().Be(2);
        warnings.Count.Should().Be(1);
        warnings[0].Should().Contain("in-memory cache");
    }

    [Fact]
    public void Token_cache_strict_mode_throws_actionable_cache_error()
    {
        var credential = new TokenCacheCredential(
            new FailingCredential("Persistence check failed: libsecret"),
            memory: null,
            _ => { });

        Action act = () => credential.GetToken(new TokenRequestContext(["scope"]), CancellationToken.None);

        act.Should().Throw<MailServiceException>()
            .Where(e => e.Code == "auth-cache-unavailable")
            .WithMessage("*FallbackToMemoryTokenCache*");
    }

    [Fact]
    public void Unencrypted_cache_warning_is_written_once()
    {
        List<string> warnings = [];
        var factory = new TokenCredentialFactory(warnings.Add);
        var options = new GraphAuthOptions
        {
            AuthMode = AuthMode.Delegated,
            DelegatedFlow = DelegatedFlow.DeviceCode,
            TenantId = "t",
            ClientId = "c",
            UnsafeAllowUnencryptedTokenCache = true
        };

        factory.GetCredential(options);
        factory.GetCredential(options);

        warnings.Count(w => w.Contains("unencrypted token-cache", StringComparison.OrdinalIgnoreCase)).Should().Be(1);
    }

    private sealed class FailingCredential(string message) : TokenCredential
    {
        public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken) =>
            throw new InvalidOperationException(message);

        public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken) =>
            ValueTask.FromException<AccessToken>(new InvalidOperationException(message));
    }

    private sealed class CountingCredential : TokenCredential
    {
        public int Calls { get; private set; }

        public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            Calls++;
            return new AccessToken("token", DateTimeOffset.UtcNow.AddMinutes(5));
        }

        public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            Calls++;
            return ValueTask.FromResult(new AccessToken("token", DateTimeOffset.UtcNow.AddMinutes(5)));
        }
    }
}
