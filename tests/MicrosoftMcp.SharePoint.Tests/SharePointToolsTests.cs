using System.Text.Json;
using AwesomeAssertions;
using Microsoft.Extensions.Logging.Abstractions;
using ModelContextProtocol.Protocol;

namespace MicrosoftMcp.SharePoint.Tests;

internal static class ToolResults
{
    private static readonly JsonSerializerOptions Read =
        new(ToolResult.Json) { PropertyNameCaseInsensitive = true };

    internal static T Ok<T>(CallToolResult result)
    {
        Assert.False(result.IsError is true);
        return JsonSerializer.Deserialize<T>(ToolResult.ReadText(result), Read)
            ?? throw new InvalidOperationException("Tool result was null JSON.");
    }

    internal static string Fail(CallToolResult result)
    {
        Assert.True(result.IsError is true);
        return ToolResult.ReadText(result);
    }
}

public sealed class SharePointToolsTests
{
    private const string Drive = "drive-eng";

    private static SharePointTools Create(IGraphSharePointService sites) =>
        new(sites, NullLogger<SharePointTools>.Instance);

    [Fact]
    public async Task Site_discovery_finds_site_and_drives()
    {
        var tools = Create(new FakeGraphSharePointService());

        var found = ToolResults.Ok<List<SharePointSiteInfo>>(await tools.sharepoint_search_sites("engineer"));
        found.Should().Contain(s => s.Id == "site-eng");

        var root = ToolResults.Ok<SharePointSiteInfo>(await tools.sharepoint_get_site("root"));
        root.Id.Should().Be("site-root");

        var drives = ToolResults.Ok<List<SharePointDriveInfo>>(await tools.sharepoint_list_site_drives("site-eng"));
        drives.Should().Contain(d => d.Id == "drive-eng");
    }

    [Fact]
    public async Task Browse_drive_lists_and_gets_items()
    {
        var tools = Create(new FakeGraphSharePointService());

        var root = ToolResults.Ok<List<DriveItemSummary>>(await tools.sharepoint_list_children(Drive, "root"));
        root.Should().Contain(i => i.Name == "Protokolle");

        var byPath = ToolResults.Ok<DriveItemSummary>(await tools.sharepoint_get_item(Drive, "/Protokolle"));
        byPath.IsFolder.Should().BeTrue();

        var found = ToolResults.Ok<List<DriveItemSummary>>(await tools.sharepoint_search_files(Drive, "notiz"));
        found.Should().Contain(i => i.Id == "f-note");
    }

    [Fact]
    public async Task Drives_are_isolated_from_each_other()
    {
        var tools = Create(new FakeGraphSharePointService());

        var eng = ToolResults.Ok<List<DriveItemSummary>>(await tools.sharepoint_list_children("drive-eng", "root"));
        var root = ToolResults.Ok<List<DriveItemSummary>>(await tools.sharepoint_list_children("drive-root", "root"));
        eng.Should().NotBeEquivalentTo(root);
        root.Should().Contain(i => i.Name == "liesmich.txt");
    }

    [Fact]
    public async Task Download_text_and_binary_with_fake()
    {
        var tools = Create(new FakeGraphSharePointService());

        var text = ToolResults.Ok<FileContentDto>(await tools.sharepoint_download_file(Drive, "f-note"));
        text.Encoding.Should().Be("text");
        text.Text.Should().Contain("Hallo Welt");

        var binary = ToolResults.Ok<FileContentDto>(await tools.sharepoint_download_file(Drive, "/Protokolle/bild.png"));
        binary.Encoding.Should().Be("base64");
        binary.DataBase64.Should().NotBeNullOrEmpty();
    }

    [Fact]
    public async Task Download_errors_carry_codes()
    {
        var tools = Create(new FakeGraphSharePointService());

        ToolResults.Fail(await tools.sharepoint_download_file(Drive, "f-big")).Should().Contain("[attachment-too-large]");
        ToolResults.Fail(await tools.sharepoint_download_file(Drive, "f-docs")).Should().Contain("[invalid-request]");
        ToolResults.Fail(await tools.sharepoint_download_file(Drive, "nope")).Should().Contain("[item-not-found]");
        ToolResults.Fail(await tools.sharepoint_download_file("drive-nope", "root")).Should().Contain("[drive-not-found]");
        ToolResults.Fail(await tools.sharepoint_download_file("", "root")).Should().Contain("[invalid-request]");
    }

    [Fact]
    public async Task Site_errors_carry_codes()
    {
        var tools = Create(new FakeGraphSharePointService());

        ToolResults.Fail(await tools.sharepoint_search_sites("")).Should().Contain("[invalid-request]");
        ToolResults.Fail(await tools.sharepoint_get_site("site-nope")).Should().Contain("[site-not-found]");
        ToolResults.Fail(await tools.sharepoint_list_site_drives("site-nope")).Should().Contain("[site-not-found]");
    }

    [Fact]
    public async Task Create_upload_move_flow_with_fake()
    {
        var tools = Create(new FakeGraphSharePointService());

        var folder = ToolResults.Ok<DriveItemSummary>(await tools.sharepoint_create_folder(Drive, "Belege"));
        folder.IsFolder.Should().BeTrue();

        var uploaded = ToolResults.Ok<DriveItemSummary>(
            await tools.sharepoint_upload_file(Drive, "notiz.txt", folder.Id, contentText: "Inhalt"));
        uploaded.Name.Should().Be("notiz.txt");

        var moved = ToolResults.Ok<DriveItemSummary>(
            await tools.sharepoint_move_item(Drive, uploaded.Id, newName: "notiz-neu.txt"));
        moved.Name.Should().Be("notiz-neu.txt");
    }
}
