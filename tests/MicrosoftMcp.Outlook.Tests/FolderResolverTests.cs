using AwesomeAssertions;
using MicrosoftMcp.Common;

namespace MicrosoftMcp.Outlook.Tests;

public sealed class FolderResolverTests
{
    private static readonly IReadOnlyList<FolderInfo> Folders =
    [
        new("id-inbox", "Inbox", 3, 1),
        new("id-custom", "Projekte", 0, 0)
    ];

    [Theory]
    [InlineData("inbox", "inbox")]
    [InlineData("Archive", "archive")]
    [InlineData("  DeletedItems ", "deleteditems")]
    [InlineData("DRAFTS", "drafts")]
    public void Wellknown_names_pass_through_normalized(string input, string expected) =>
        FolderResolver.Resolve(input, Folders).Should().Be(expected);

    [Fact]
    public void Resolves_by_id_case_insensitive() =>
        FolderResolver.Resolve("ID-CUSTOM", Folders).Should().Be("id-custom");

    [Fact]
    public void Resolves_by_display_name_case_insensitive() =>
        FolderResolver.Resolve("projekte", Folders).Should().Be("id-custom");

    [Fact]
    public void Unknown_folder_throws_coded_error_with_hint()
    {
        var ex = Assert.Throws<MailServiceException>(() => FolderResolver.Resolve("gibts-nicht", Folders));
        ex.Code.Should().Be("folder-not-found");
        ex.Message.Should().Contain("Next:");
    }

    [Fact]
    public void Blank_destination_throws() =>
        Assert.Throws<ArgumentException>(() => FolderResolver.Resolve("  ", Folders));
}
