using AwesomeAssertions;
using MicrosoftMcp.Common;
using MicrosoftMcp.Host.Setup;

namespace MicrosoftMcp.Host.Tests;

/// <summary>SharePoint/Planner run on Sites/Group APIs without personal-account
/// support: setup must reject personal, doctor must fail a consumers tenant.</summary>
public sealed class AccountValidationTests
{
    [Fact]
    public void Personal_is_rejected_for_sharepoint_and_planner()
    {
        SetupGuide.ValidateAccountCombination(["sharepoint"], "personal").Should().NotBeNull();
        SetupGuide.ValidateAccountCombination(["planner"], "personal").Should().NotBeNull();
        SetupGuide.ValidateAccountCombination(["outlook", "sharepoint"], "personal").Should().NotBeNull();
    }

    [Fact]
    public void Work_and_personal_without_sharepoint_or_planner_are_valid()
    {
        SetupGuide.ValidateAccountCombination(["sharepoint", "planner"], "work").Should().BeNull();
        SetupGuide.ValidateAccountCombination(["outlook", "calendar"], "personal").Should().BeNull();
        SetupGuide.ValidateAccountCombination(["todo"], "personal").Should().BeNull();
    }

    private static GraphAuthOptions Options(string tenant)
    {
        return new GraphAuthOptions
        {
            AuthMode = AuthMode.Delegated,
            TenantId = tenant,
            ClientId = "11111111-1111-1111-1111-111111111111"
        };
    }

    [Fact]
    public void Doctor_fails_consumers_tenant_for_sharepoint()
    {
        var checks = DoctorChecks.Run(["sharepoint"], Options("consumers"), policyPath: null, policyError: null);

        var account = checks.First(c => c.Id == "account");
        account.Ok.Should().BeFalse();
        account.Next.Should().Contain("Graph__TenantId");
    }

    [Fact]
    public void Doctor_fails_consumers_tenant_for_planner()
    {
        var checks = DoctorChecks.Run(["planner"], Options("consumers"), policyPath: null, policyError: null);

        checks.First(c => c.Id == "account").Ok.Should().BeFalse();
    }

    [Fact]
    public void Doctor_fails_account_check_for_missing_or_invalid_tenant()
    {
        foreach (string tenant in new[] { "", "YOUR-TENANT-ID", "not-a-tenant" })
        {
            var account = DoctorChecks.Run(["sharepoint"], Options(tenant), policyPath: null, policyError: null)
                .First(c => c.Id == "account");

            account.Ok.Should().BeFalse($"tenant '{tenant}' cannot establish an org account");
            account.Next.Should().Contain("Graph__TenantId");
        }
    }

    [Fact]
    public void Doctor_passes_org_tenant_for_sharepoint_and_planner()
    {
        var checks = DoctorChecks.Run(
            ["sharepoint", "planner"],
            Options("11111111-2222-3333-4444-555555555555"),
            policyPath: null,
            policyError: null);

        checks.Should().OnlyContain(c => c.Ok);
    }

    [Fact]
    public void Doctor_warns_for_common_tenant_with_sharepoint()
    {
        var account = DoctorChecks.Run(["sharepoint"], Options("common"), policyPath: null, policyError: null)
            .First(c => c.Id == "account");

        account.Ok.Should().BeTrue();
        account.Next.Should().Contain("work account");
    }

    [Fact]
    public void Doctor_passes_organizations_tenant_for_planner()
    {
        var account = DoctorChecks.Run(["planner"], Options("organizations"), policyPath: null, policyError: null)
            .First(c => c.Id == "account");

        account.Ok.Should().BeTrue();
    }

    [Fact]
    public void Doctor_has_no_account_check_without_sharepoint_or_planner()
    {
        var checks = DoctorChecks.Run(["outlook", "calendar"], Options("consumers"), policyPath: null, policyError: null);

        checks.Should().NotContain(c => c.Id == "account");
    }
}
