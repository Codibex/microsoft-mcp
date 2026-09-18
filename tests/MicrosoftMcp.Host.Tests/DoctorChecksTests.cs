using AwesomeAssertions;
using MicrosoftMcp.Common;
using MicrosoftMcp.Host.Setup;

namespace MicrosoftMcp.Host.Tests;

/// <summary>Offline doctor: config validation without network and without leaking secrets.</summary>
public sealed class DoctorChecksTests
{
    private static GraphAuthOptions Delegated(string tenant = "common", string client = "11111111-1111-1111-1111-111111111111")
    {
        return new GraphAuthOptions { AuthMode = AuthMode.Delegated, TenantId = tenant, ClientId = client };
    }

    [Fact]
    public void Missing_ids_fail_with_next_hint()
    {
        var checks = DoctorChecks.Run(["outlook"], new GraphAuthOptions(), policyPath: null, policyError: null);

        checks.First(c => c.Id == "tenant").Ok.Should().BeFalse();
        checks.First(c => c.Id == "client").Ok.Should().BeFalse();
        checks.First(c => c.Id == "tenant").Next.Should().Contain("Graph__TenantId");
    }

    [Fact]
    public void Placeholders_fail()
    {
        var options = Delegated(tenant: "YOUR-TENANT-ID", client: "YOUR-CLIENT-ID");
        var checks = DoctorChecks.Run(["calendar"], options, policyPath: null, policyError: null);

        checks.First(c => c.Id == "tenant").Ok.Should().BeFalse();
        checks.First(c => c.Id == "client").Ok.Should().BeFalse();
    }

    [Fact]
    public void Valid_delegated_config_passes()
    {
        var checks = DoctorChecks.Run(["outlook", "calendar"], Delegated(), policyPath: null, policyError: null);

        checks.Should().OnlyContain(c => c.Ok);
    }

    [Fact]
    public void Delegated_cache_uses_memory_fallback_by_default()
    {
        var cache = DoctorChecks.Run(["calendar"], Delegated(), policyPath: null, policyError: null)
            .First(c => c.Id == "cache");

        cache.Ok.Should().BeTrue();
        cache.Message.Should().Contain("in-memory fallback");
    }

    [Fact]
    public void Unencrypted_cache_is_explicitly_reported()
    {
        var options = Delegated();
        options.UnsafeAllowUnencryptedTokenCache = true;

        var cache = DoctorChecks.Run(["calendar"], options, policyPath: null, policyError: null)
            .First(c => c.Id == "cache");

        cache.Ok.Should().BeTrue();
        cache.Message.Should().Contain("unencrypted");
        cache.Next.Should().Contain("Graph__EnableTokenCache=false");
    }

    [Fact]
    public void Apponly_cache_never_reports_delegated_login_requirement()
    {
        var options = Delegated();
        options.AuthMode = AuthMode.AppOnly;
        options.EnableTokenCache = false;
        options.UserIdOrUpn = "mailbox@example.com";
        options.ClientSecret = "secret";

        var cache = DoctorChecks.Run(
                ["outlook"],
                options,
                policyPath: null,
                policyError: null,
                secretServiceAvailable: () => false)
            .First(c => c.Id == "cache");

        cache.Ok.Should().BeTrue();
        cache.Message.Should().Contain("not used for AppOnly");
        cache.Message.Should().NotContain("new login");
    }

    [Fact]
    public void Strict_delegated_cache_fails_when_secret_service_is_unavailable()
    {
        if (!OperatingSystem.IsLinux())
        {
            return;
        }

        var options = Delegated();
        options.FallbackToMemoryTokenCache = false;

        var cache = DoctorChecks.Run(
                ["calendar"],
                options,
                policyPath: null,
                policyError: null,
                secretServiceAvailable: () => false)
            .First(c => c.Id == "cache");

        cache.Ok.Should().BeFalse();
        cache.Message.Should().Contain("Secret Service");
        cache.Next.Should().Contain("FallbackToMemoryTokenCache=true");
    }

    [Fact]
    public void Bad_tenant_format_fails()
    {
        var checks = DoctorChecks.Run(["outlook"], Delegated(tenant: "not-a-tenant"), policyPath: null, policyError: null);

        checks.First(c => c.Id == "tenant").Ok.Should().BeFalse();
    }

    [Fact]
    public void Apponly_with_teams_fails()
    {
        var options = Delegated();
        options.AuthMode = AuthMode.AppOnly;
        options.UserIdOrUpn = "mailbox@example.com";
        options.ClientSecret = "secret";

        var checks = DoctorChecks.Run(["teams"], options, policyPath: null, policyError: null);

        checks.First(c => c.Id == "authmode").Ok.Should().BeFalse();
    }

    [Fact]
    public void Apponly_requires_mailbox_and_secret()
    {
        var options = Delegated();
        options.AuthMode = AuthMode.AppOnly;

        var checks = DoctorChecks.Run(["outlook"], options, policyPath: null, policyError: null);

        checks.First(c => c.Id == "mailbox").Ok.Should().BeFalse();
        checks.First(c => c.Id == "secret").Ok.Should().BeFalse();
    }

    [Fact]
    public void Explicit_scopes_missing_required_fail()
    {
        var options = Delegated();
        options.DelegatedScopes = ["Mail.Read"];

        var checks = DoctorChecks.Run(["outlook"], options, policyPath: null, policyError: null);

        checks.First(c => c.Id == "scopes").Ok.Should().BeFalse();
    }

    [Fact]
    public void Missing_policy_is_ok_with_hint_for_messaging_servers()
    {
        var checks = DoctorChecks.Run(["outlook"], Delegated(), policyPath: null, policyError: null);

        var policy = checks.First(c => c.Id == "policy");
        policy.Ok.Should().BeTrue();
        policy.Next.Should().Contain("policy.json");
    }

    [Fact]
    public void Unreadable_policy_fails()
    {
        var checks = DoctorChecks.Run(["outlook"], Delegated(), policyPath: null, policyError: "not readable");

        checks.First(c => c.Id == "policy").Ok.Should().BeFalse();
    }
}
