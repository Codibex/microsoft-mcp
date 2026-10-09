using AwesomeAssertions;
using MicrosoftMcp.Host.Setup;

namespace MicrosoftMcp.Host.Tests;

/// <summary>Bare --help/--version must print instead of starting a stdio server.</summary>
public sealed class HostHelpTests
{
    [Theory]
    [InlineData("--help")]
    [InlineData("-h")]
    [InlineData("help")]
    public void Help_flags_are_detected(string flag)
    {
        HostHelp.ShouldShowHelp([flag]).Should().BeTrue();
        HostHelp.ShouldShowHelp(["--servers", "outlook", flag]).Should().BeTrue();
    }

    [Fact]
    public void Version_flag_is_detected()
    {
        HostHelp.ShouldShowVersion(["--version"]).Should().BeTrue();
        HostHelp.ShouldShowVersion(["--servers", "outlook"]).Should().BeFalse();
    }

    [Fact]
    public void Plain_server_args_show_neither_help_nor_version()
    {
        HostHelp.ShouldShowHelp([]).Should().BeFalse();
        HostHelp.ShouldShowHelp(["--servers", "outlook"]).Should().BeFalse();
        HostHelp.ShouldShowVersion([]).Should().BeFalse();
    }

    [Fact]
    public void Usage_lists_the_offline_commands()
    {
        string usage = HostHelp.UsageText();

        usage.Should().Contain("setup");
        usage.Should().Contain("doctor");
        usage.Should().Contain("policy migrate");
    }

    [Fact]
    public void Version_names_the_binary()
    {
        HostHelp.VersionText().Should().StartWith("microsoft-mcp ");
    }
}
