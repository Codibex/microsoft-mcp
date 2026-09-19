using AwesomeAssertions;
using MicrosoftMcp.Common;

namespace MicrosoftMcp.Outlook.Tests;

public sealed class FolderResolverTests
{
    private static readonly IReadOnlyList<FolderInfo> Folders =
    [
        new("id-inbox", "Inbox", 3, 1, "root", "Inbox"),
        new("id-custom", "Projekte", 0, 0, "root", "Projekte"),
        new("id-school", "School", 0, 0, "root", "School"),
        new("id-kids", "Kids", 0, 0, "id-school", "School/Kids"),
        new("id-sophie", "Sophie", 0, 0, "id-kids", "School/Kids/Sophie")
    ];

    [Theory]
    [InlineData("inbox", "inbox")]
    [InlineData("Archive", "archive")]
    [InlineData("  DeletedItems ", "deleteditems")]
    [InlineData("DRAFTS", "drafts")]
    [InlineData("SyncIssues", "syncissues")]
    [InlineData("RecoverableItemsDeletions", "recoverableitemsdeletions")]
    public void Wellknown_names_pass_through_normalized(string input, string expected) =>
        FolderResolver.Resolve(input, Folders).Should().Be(expected);

    [Fact]
    public void Resolves_by_id_case_insensitive() =>
        FolderResolver.Resolve("ID-CUSTOM", Folders).Should().Be("id-custom");

    [Fact]
    public void Resolves_by_display_name_case_insensitive() =>
        FolderResolver.Resolve("projekte", Folders).Should().Be("id-custom");

    [Fact]
    public void Four_value_folder_constructor_and_deconstructor_remain_usable()
    {
        var folder = new FolderInfo("id", "Name", 3, 1);

        folder.ParentId.Should().BeNull();
        folder.Path.Should().BeNull();
        var (id, displayName, totalCount, unreadCount) = folder;
        (id, displayName, totalCount, unreadCount).Should().Be(("id", "Name", 3, 1));
    }

    [Fact]
    public void Resolves_by_nested_path_case_insensitive() =>
        FolderResolver.Resolve(" school/kids/sophie ", Folders).Should().Be("id-sophie");

    [Fact]
    public void Ambiguous_display_name_is_rejected()
    {
        var folders = new[]
        {
            new FolderInfo("id-one", "Kunden", 0, 0, "root", "Privat/Kunden"),
            new FolderInfo("id-two", "Kunden", 0, 0, "root", "Arbeit/Kunden")
        };

        var ex = Assert.Throws<GraphServiceException>(() => FolderResolver.Resolve("Kunden", folders));
        ex.Code.Should().Be("invalid-request");
        ex.Message.Should().Contain("full folder path");
    }

    [Fact]
    public void Unknown_folder_throws_coded_error_with_hint()
    {
        var ex = Assert.Throws<GraphServiceException>(() => FolderResolver.Resolve("gibts-nicht", Folders));
        ex.Code.Should().Be("folder-not-found");
        ex.Message.Should().Contain("Next:");
    }

    [Fact]
    public void Blank_destination_throws() =>
        Assert.Throws<ArgumentException>(() => FolderResolver.Resolve("  ", Folders));
}
