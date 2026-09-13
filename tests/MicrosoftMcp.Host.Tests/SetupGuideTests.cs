using AwesomeAssertions;
using MicrosoftMcp.Common;
using MicrosoftMcp.Host.Setup;

namespace MicrosoftMcp.Host.Tests;

/// <summary>Setup-Wizard: Kombinationen, Scopes, Tenant-Hints, mcp.json-Snippet.</summary>
public sealed class SetupGuideTests
{
    [Fact]
    public void AppOnly_is_rejected_for_non_outlook_servers()
    {
        SetupGuide.ValidateCombination(["outlook", "calendar"], AuthMode.AppOnly).Should().NotBeNull();
        SetupGuide.ValidateCombination(["teams"], AuthMode.AppOnly).Should().NotBeNull();
        SetupGuide.ValidateCombination(["onedrive"], AuthMode.AppOnly).Should().NotBeNull();
    }

    [Fact]
    public void Outlook_only_apponly_and_any_delegated_are_valid()
    {
        SetupGuide.ValidateCombination(["outlook"], AuthMode.AppOnly).Should().BeNull();
        SetupGuide.ValidateCombination(["outlook", "calendar"], AuthMode.Delegated).Should().BeNull();
    }

    [Fact]
    public void Scopes_follow_the_selection()
    {
        SetupGuide.ScopesFor(["outlook", "calendar"])
            .Should().Equal("Mail.Read", "Mail.ReadWrite", "Calendars.Read");
    }

    [Fact]
    public void Tenant_hint_depends_on_account()
    {
        SetupGuide.TenantHint("personal").Should().Be("common");
        SetupGuide.TenantHint("work").Should().Be("<tenant-guid>");
    }

    [Fact]
    public void Claude_uses_mcpServers_vscode_uses_servers()
    {
        SetupGuide.TopLevelKey(SetupClient.Claude).Should().Be("mcpServers");
        SetupGuide.TopLevelKey(SetupClient.Vscode).Should().Be("servers");
        SetupGuide.TopLevelKey(SetupClient.Generic).Should().Be("servers");
    }

    [Fact]
    public void Snippet_contains_binary_servers_and_placeholders()
    {
        string json = SetupGuide.BuildMcpJson(
            SetupClient.Vscode, "/opt/microsoft-mcp/microsoft-mcp",
            ["outlook", "calendar"], "common", AuthMode.Delegated, headless: false);

        json.Should().Contain("\"servers\"");
        json.Should().Contain("/opt/microsoft-mcp/microsoft-mcp");
        json.Should().Contain("outlook,calendar");
        json.Should().Contain("Graph__TenantId");
        json.Should().NotContain("ClientSecret");
    }

    [Fact]
    public void Snippet_adds_apponly_and_headless_env()
    {
        string json = SetupGuide.BuildMcpJson(
            SetupClient.Claude, "/bin/microsoft-mcp",
            ["outlook"], "<tenant-guid>", AuthMode.AppOnly, headless: true);

        json.Should().Contain("\"mcpServers\"");
        json.Should().Contain("Graph__AuthMode");
        json.Should().Contain("Graph__ClientSecret");
        json.Should().Contain("DeviceCode");
    }
}
