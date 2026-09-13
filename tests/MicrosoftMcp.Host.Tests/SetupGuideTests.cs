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
        SetupGuide.TopLevelKey(SetupClient.Opencode).Should().Be("mcp");
        SetupGuide.TopLevelKey(SetupClient.Codex).Should().Be("mcp_servers");
        SetupGuide.TopLevelKey(SetupClient.Openclaw).Should().Be("mcp.servers");
        SetupGuide.TopLevelKey(SetupClient.Hermes).Should().Be("mcp_servers");
    }

    [Fact]
    public void Client_parsing_supports_all_clients()
    {
        SetupGuide.ParseClient("vscode").Should().Be(SetupClient.Vscode);
        SetupGuide.ParseClient("claude-desktop").Should().Be(SetupClient.Claude);
        SetupGuide.ParseClient("opencode").Should().Be(SetupClient.Opencode);
        SetupGuide.ParseClient("codex").Should().Be(SetupClient.Codex);
        SetupGuide.ParseClient("openclaw").Should().Be(SetupClient.Openclaw);
        SetupGuide.ParseClient("hermes-agent").Should().Be(SetupClient.Hermes);
        SetupGuide.ParseClient("unknown").Should().Be(SetupClient.Generic);
    }

    [Fact]
    public void Config_files_point_to_the_right_place()
    {
        SetupGuide.ConfigFile(SetupClient.Vscode).Should().Be(".vscode/mcp.json");
        SetupGuide.ConfigFile(SetupClient.Codex).Should().Be("~/.codex/config.toml");
        SetupGuide.ConfigFile(SetupClient.Hermes).Should().Be("~/.hermes/config.yaml");
        SetupGuide.Format(SetupClient.Codex).Should().Be(SnippetFormat.Toml);
        SetupGuide.Format(SetupClient.Hermes).Should().Be(SnippetFormat.Yaml);
        SetupGuide.Format(SetupClient.Vscode).Should().Be(SnippetFormat.Json);
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

    [Fact]
    public void Opencode_snippet_uses_command_array_and_environment()
    {
        string json = SetupGuide.BuildMcpJson(
            SetupClient.Opencode, "/opt/microsoft-mcp/microsoft-mcp",
            ["outlook"], "common", AuthMode.Delegated, headless: false);

        json.Should().Contain("\"mcp\"");
        json.Should().Contain("\"type\": \"local\"");
        json.Should().Contain("\"environment\"");
        json.Should().Contain("[\"/opt/microsoft-mcp/microsoft-mcp\", \"--servers\", \"outlook\"]");
    }

    [Fact]
    public void Codex_snippet_is_toml_with_env_table()
    {
        string toml = SetupGuide.BuildMcpJson(
            SetupClient.Codex, "/opt/microsoft-mcp/microsoft-mcp",
            ["calendar"], "common", AuthMode.Delegated, headless: false);

        toml.Should().Contain("[mcp_servers.m365]");
        toml.Should().Contain("command = \"/opt/microsoft-mcp/microsoft-mcp\"");
        toml.Should().Contain("[mcp_servers.m365.env]");
        toml.Should().Contain("Graph__TenantId = \"common\"");
    }

    [Fact]
    public void Hermes_snippet_is_yaml_with_mcp_servers()
    {
        string yaml = SetupGuide.BuildMcpJson(
            SetupClient.Hermes, "/opt/microsoft-mcp/microsoft-mcp",
            ["outlook"], "common", AuthMode.Delegated, headless: false);

        yaml.Should().Contain("mcp_servers:");
        yaml.Should().Contain("command: \"/opt/microsoft-mcp/microsoft-mcp\"");
        yaml.Should().Contain("Graph__ClientId");
    }

    [Fact]
    public void Openclaw_snippet_nests_mcp_servers_with_enabled_flag()
    {
        string json = SetupGuide.BuildMcpJson(
            SetupClient.Openclaw, "/opt/microsoft-mcp/microsoft-mcp",
            ["teams"], "common", AuthMode.Delegated, headless: false);

        json.Should().Contain("\"mcp\"");
        json.Should().Contain("\"servers\"");
        json.Should().Contain("\"enabled\": true");
    }
}
