using AwesomeAssertions;
using MicrosoftMcp.Common;

namespace MicrosoftMcp.Common.Tests;

public sealed class ServerSelectionTests
{
    [Fact]
    public void Empty_selects_all_in_stable_order() =>
        ServerSelection.Parse(null).Should().Equal("outlook", "onedrive", "calendar", "teams", "sharepoint");

    [Fact]
    public void Parses_and_normalizes_selection() =>
        ServerSelection.Parse("teams, outlook").Should().Equal("outlook", "teams");

    [Fact]
    public void Unknown_names_throw_with_valid_list()
    {
        Action act = () => ServerSelection.Parse("outlook,mail");
        act.Should().Throw<InvalidOperationException>().WithMessage("*outlook*onedrive*calendar*teams*");
    }

    [Fact]
    public void Scopes_follow_selection_without_duplicates()
    {
        ServerSelection.DefaultScopesFor(["teams"]).Should().Contain("Chat.Read");
        ServerSelection.DefaultScopesFor(["outlook"]).Should()
            .NotContain(s => s.StartsWith("Files", StringComparison.Ordinal));
        ServerSelection.DefaultScopesFor(ServerSelection.All).Should().HaveCount(2 + 3 + 2 + 9 + 1);
    }

    [Fact]
    public void Sharepoint_scopes_cover_site_discovery_and_files()
    {
        ServerSelection.DefaultScopesFor(["sharepoint"]).Should()
            .BeEquivalentTo("Sites.Read.All", "Files.Read", "Files.ReadWrite");
        ServerSelection.DefaultScopesFor(["teams", "sharepoint"]).Should().Contain("Channel.ReadBasic.All");
        // Least-privileged delegated scope for the Teams filesFolder bridge (work accounts).
        ServerSelection.DefaultScopesFor(["teams"]).Should().Contain("Files.Read.All");
    }
}
