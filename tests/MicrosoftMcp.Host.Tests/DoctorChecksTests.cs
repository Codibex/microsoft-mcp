using AwesomeAssertions;
using MicrosoftMcp.Common;
using MicrosoftMcp.Host.Setup;

namespace MicrosoftMcp.Host.Tests;

/// <summary>Offline-Doctor: Config-Validierung ohne Netzwerk und ohne Secrets-Abfluss.</summary>
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
