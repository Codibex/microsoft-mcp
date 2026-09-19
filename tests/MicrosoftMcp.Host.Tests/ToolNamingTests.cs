using System.Reflection;
using AwesomeAssertions;
using MicrosoftMcp.Calendar;
using MicrosoftMcp.OneDrive;
using MicrosoftMcp.Outlook;
using MicrosoftMcp.Teams;
using ModelContextProtocol.Server;

namespace MicrosoftMcp.Host.Tests;

/// <summary>
/// Enforces the MCP tool-naming convention (spec: names SHOULD be unique
/// within a server; multi-domain collisions are disambiguated by a domain
/// prefix): every tool is {domain}_{verb}, unique across all toolsets.
/// </summary>
public sealed class ToolNamingTests
{
    private static readonly (Type Tools, string Prefix)[] ToolSets =
    [
        (typeof(OutlookTools), "outlook_"),
        (typeof(DriveTools), "onedrive_"),
        (typeof(CalendarTools), "calendar_"),
        (typeof(TeamsTools), "teams_")
    ];

    private static IReadOnlyList<(string ToolSet, string Name)> AllTools() =>
        [.. ToolSets.SelectMany(t => t.Tools
            .GetMethods(BindingFlags.Public | BindingFlags.Instance)
            .Where(m => m.GetCustomAttribute<McpServerToolAttribute>() is not null)
            .Select(m => (t.Tools.Name, m.Name)))];

    [Fact]
    public void All_tool_names_are_unique_across_domains()
    {
        var names = AllTools().Select(t => t.Name).ToList();
        names.Should().HaveCount(46);
        names.Should().OnlyHaveUniqueItems();
    }

    [Fact]
    public void Every_tool_carries_its_domain_prefix()
    {
        foreach (var (toolSet, name) in AllTools())
        {
            string expected = ToolSets.First(t => t.Tools.Name == toolSet).Prefix;
            name.Should().StartWith(expected, $"tool {name} in {toolSet}");
        }
    }

    [Fact]
    public void Tool_names_use_only_spec_allowed_characters()
    {
        foreach (var (_, name) in AllTools())
        {
            name.Should().MatchRegex("^[a-z0-9_]+$");
            name.Length.Should().BeInRange(1, 128);
        }
    }
}
