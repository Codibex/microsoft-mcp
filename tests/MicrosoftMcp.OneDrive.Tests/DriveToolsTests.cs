using System.Text.Json;
using AwesomeAssertions;
using Microsoft.Extensions.Logging.Abstractions;
using ModelContextProtocol.Protocol;
using NSubstitute;

namespace MicrosoftMcp.OneDrive.Tests;

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

public sealed class DriveToolsTests
{
    private static DriveTools Create(IGraphDriveService drive) =>
        new(drive, NullLogger<DriveTools>.Instance);

    [Fact]
    public async Task Browse_drive_lists_and_gets_items()
    {
        var tools = Create(new FakeGraphDriveService());

        var drive = ToolResults.Ok<DriveInfoDto>(await tools.get_drive());
        drive.Name.Should().Be("OneDrive");

        ToolResults.Ok<List<DriveInfoDto>>(await tools.list_drives()).Should().HaveCount(1);

        var root = ToolResults.Ok<List<DriveItemSummary>>(await tools.list_children("root"));
        root.Should().Contain(i => i.Name == "Dokumente");

        var byPath = ToolResults.Ok<DriveItemSummary>(await tools.get_item("/Dokumente"));
        byPath.IsFolder.Should().BeTrue();

        var found = ToolResults.Ok<List<DriveItemSummary>>(await tools.search_files("notiz"));
        found.Should().Contain(i => i.Id == "f-note");
    }

    [Fact]
    public async Task Download_text_and_binary_with_fake()
    {
        var tools = Create(new FakeGraphDriveService());

        var text = ToolResults.Ok<FileContentDto>(await tools.download_file("f-note"));
        text.Encoding.Should().Be("text");
        text.Text.Should().Contain("Hallo Welt");

        var binary = ToolResults.Ok<FileContentDto>(await tools.download_file("/Dokumente/bild.png"));
        binary.Encoding.Should().Be("base64");
        binary.DataBase64.Should().NotBeNullOrEmpty();
    }

    [Fact]
    public async Task Download_errors_carry_codes()
    {
        var tools = Create(new FakeGraphDriveService());

        ToolResults.Fail(await tools.download_file("f-big")).Should().Contain("[attachment-too-large]");
        ToolResults.Fail(await tools.download_file("f-docs")).Should().Contain("[invalid-request]");
        ToolResults.Fail(await tools.download_file("nope")).Should().Contain("[item-not-found]");
    }

    [Fact]
    public async Task Create_upload_move_flow_with_fake()
    {
        var tools = Create(new FakeGraphDriveService());

        var folder = ToolResults.Ok<DriveItemSummary>(await tools.create_folder("Belege"));
        folder.IsFolder.Should().BeTrue();

        var uploaded = ToolResults.Ok<DriveItemSummary>(
            await tools.upload_file("neu.txt", folder.Id, contentText: "Inhalt"));
        uploaded.Name.Should().Be("neu.txt");

        var downloaded = ToolResults.Ok<FileContentDto>(await tools.download_file(uploaded.Id));
        downloaded.Text.Should().Be("Inhalt");

        var moved = ToolResults.Ok<DriveItemSummary>(
            await tools.move_item(uploaded.Id, newParentRef: "root", newName: "umbenannt.txt"));
        moved.Name.Should().Be("umbenannt.txt");
        moved.ParentId.Should().Be("root");
    }

    [Fact]
    public async Task Upload_validation_errors_carry_codes()
    {
        var tools = Create(new FakeGraphDriveService());

        ToolResults.Fail(await tools.upload_file("x.txt", contentBase64: "!!!"))
            .Should().Contain("[invalid-request]");
        ToolResults.Fail(await tools.upload_file("x.txt"))
            .Should().Contain("[invalid-request]");
        ToolResults.Fail(await tools.move_item("f-note"))
            .Should().Contain("[invalid-request]");
    }

    [Fact]
    public async Task Unexpected_backend_failures_are_mapped_not_leaked()
    {
        IGraphDriveService failing = Substitute.For<IGraphDriveService>();
        failing.GetDriveAsync(Arg.Any<CancellationToken>())
            .Returns<Task<DriveInfoDto>>(_ => throw new HttpRequestException("no route"));
        var tools = Create(failing);

        ToolResults.Fail(await tools.get_drive()).Should().Contain("[service-unavailable]");
    }
}
