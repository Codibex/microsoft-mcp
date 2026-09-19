using System.Text.Json;
using AwesomeAssertions;
using Microsoft.Extensions.Logging.Abstractions;
using ModelContextProtocol.Protocol;
using NSubstitute;

namespace MicrosoftMcp.Outlook.Tests;

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

public sealed class OutlookToolsTests
{
    private static OutlookTools Create(IGraphMailService mail) =>
        new(mail, NullLogger<OutlookTools>.Instance);
    [Fact]
    public async Task Search_delegates_query_and_returns_results()
    {
        IGraphMailService mail = Substitute.For<IGraphMailService>();
        mail.SearchAsync(Arg.Any<EmailQuery>(), Arg.Any<CancellationToken>())
            .Returns([new EmailSummary("m1", "Hi", null, [], null, false, false, [], null, null)]);
        var tools = Create(mail);

        var result = ToolResults.Ok<List<EmailSummary>>(await tools.outlook_search_emails("Hi", top: 10));

        result.Should().HaveCount(1);
        result[0].Id.Should().Be("m1");
    }

    [Fact]
    public async Task Search_passes_folder_and_top_through_with_fake()
    {
        var fake = new FakeGraphMailService();
        var tools = Create(fake);

        await tools.outlook_search_emails("Rechnung", folder: "Inbox", top: 5);

        fake.LastQuery.Should().NotBeNull();
        fake.LastQuery!.Query.Should().Be("Rechnung");
        fake.LastQuery.Folder.Should().Be("Inbox");
        fake.LastQuery.Top.Should().Be(5);
    }

    [Fact]
    public async Task Triage_flow_move_archive_and_trash_with_fake()
    {
        var tools = Create(new FakeGraphMailService());

        ToolResults.Ok<EmailSummary>(await tools.outlook_move_email("m1", "Projekte")).Id.Should().Be("m1");
        ToolResults.Ok<EmailDetail>(await tools.outlook_read_email("m1")).Should().NotBeNull();

        var inbox = ToolResults.Ok<List<EmailSummary>>(await tools.outlook_search_emails(folder: "inbox"));
        inbox.Should().NotContain(m => m.Id == "m1");

        await tools.outlook_archive_email("m2");
        ToolResults.Ok<List<EmailSummary>>(await tools.outlook_search_emails(folder: "archive"))
            .Should().Contain(m => m.Id == "m2");

        ToolResults.Ok<string>(await tools.outlook_delete_email("m2")).Should().Contain("trash");
        ToolResults.Ok<List<EmailSummary>>(await tools.outlook_search_emails(folder: "deleteditems"))
            .Should().Contain(m => m.Id == "m2");
    }

    [Fact]
    public async Task Draft_flow_create_reply_forward_update_with_fake()
    {
        var tools = Create(new FakeGraphMailService());

        var draft = ToolResults.Ok<EmailDetail>(await tools.outlook_create_draft(["a@x.y"], "Betreff", "Hallo"));
        draft.Id.Should().StartWith("draft-");

        var reply = ToolResults.Ok<EmailDetail>(await tools.outlook_create_reply_draft("m1", "Bin dabei"));
        reply.Subject.Should().StartWith("Re:");

        var fwd = ToolResults.Ok<EmailDetail>(await tools.outlook_create_forward_draft("m1", ["b@x.y"], "FYI"));
        fwd.Subject.Should().StartWith("Fwd:");

        var updated = ToolResults.Ok<EmailDetail>(await tools.outlook_update_draft(draft.Id, subject: "Neu"));
        updated.Subject.Should().Be("Neu");
    }

    [Fact]
    public async Task Fake_rejects_updates_to_non_drafts()
    {
        var tools = Create(new FakeGraphMailService());

        ToolResults.Fail(await tools.outlook_update_draft("m1", subject: "Nope"))
            .Should().Contain("[invalid-request]");
    }

    [Fact]
    public async Task Fake_sender_filter_is_exact()
    {
        var tools = Create(new FakeGraphMailService());

        ToolResults.Ok<List<EmailSummary>>(
                await tools.outlook_search_emails(from: "boss@example.com"))
            .Should().Contain(message => message.Id == "m1");
        ToolResults.Ok<List<EmailSummary>>(
                await tools.outlook_search_emails(from: "boss@example"))
            .Should().BeEmpty();
    }

    [Fact]
    public async Task Label_and_read_flow_with_fake()
    {
        var tools = Create(new FakeGraphMailService());

        var labeled = ToolResults.Ok<EmailDetail>(await tools.outlook_set_categories("m1", ["Blue category"], []));
        labeled.Categories.Should().Contain("Blue category");

        ToolResults.Ok<List<CategoryInfo>>(await tools.outlook_list_categories()).Should().HaveCount(2);
        ToolResults.Ok<List<AttachmentInfo>>(await tools.outlook_list_attachments("m1")).Should().HaveCount(5);
        ToolResults.Ok<List<AttachmentInfo>>(await tools.outlook_list_attachments("m2")).Should().BeEmpty();

        ToolResults.Ok<string>(await tools.outlook_mark_read("m1", true)).Should().Contain("read");
        ToolResults.Ok<string>(await tools.outlook_set_importance("m1", "high")).Should().Contain("high");
        ToolResults.Fail(await tools.outlook_set_importance("m1", "urgent")).Should().Contain("[invalid-request]");
    }

    [Fact]
    public async Task Read_attachment_text_base64_nested_reference_with_fake()
    {
        var tools = Create(new FakeGraphMailService());

        var text = ToolResults.Ok<AttachmentContent>(await tools.outlook_read_attachment("m1", "a1"));
        text.Encoding.Should().Be("text");
        text.Text.Should().Contain("42,00");
        text.Truncated.Should().BeFalse();

        var binary = ToolResults.Ok<AttachmentContent>(await tools.outlook_read_attachment("m1", "a2"));
        binary.Encoding.Should().Be("base64");
        binary.DataBase64.Should().NotBeNullOrEmpty();

        var nested = ToolResults.Ok<AttachmentContent>(await tools.outlook_read_attachment("m1", "a4"));
        nested.Encoding.Should().Be("nested");

        var reference = ToolResults.Ok<AttachmentContent>(await tools.outlook_read_attachment("m1", "a5"));
        reference.Encoding.Should().Be("reference");
        reference.SourceUrl.Should().Contain("sharepoint");
    }

    [Fact]
    public async Task Read_attachment_errors_carry_codes()
    {
        var tools = Create(new FakeGraphMailService());

        ToolResults.Fail(await tools.outlook_read_attachment("m1", "a3")).Should().Contain("[attachment-too-large]");
        ToolResults.Fail(await tools.outlook_read_attachment("m1", "nope")).Should().Contain("[attachment-not-found]");
        ToolResults.Fail(await tools.outlook_read_attachment("m1", "  ")).Should().Contain("[invalid-request]");
    }

    [Fact]
    public async Task Create_folder_then_move_with_fake()
    {
        var tools = Create(new FakeGraphMailService());

        var folder = ToolResults.Ok<FolderInfo>(await tools.outlook_create_folder("Kunden", null));
        folder.DisplayName.Should().Be("Kunden");

        await tools.outlook_move_email("m1", "Kunden");
        ToolResults.Ok<List<EmailSummary>>(await tools.outlook_search_emails(folder: "Kunden"))
            .Should().Contain(m => m.Id == "m1");
    }

    [Fact]
    public async Task Create_folder_rejects_unknown_parent_with_fake()
    {
        var tools = Create(new FakeGraphMailService());

        ToolResults.Fail(await tools.outlook_create_folder("Child", "missing"))
            .Should().Contain("[folder-not-found]");
    }

    [Fact]
    public async Task Nested_folder_path_lists_searches_and_moves_with_fake()
    {
        var tools = Create(new FakeGraphMailService());

        var folders = ToolResults.Ok<List<FolderInfo>>(await tools.outlook_list_folders());
        folders.Single(folder => folder.Id == "f-sophie").Path.Should().Be("School/Kids/Sophie");

        ToolResults.Ok<List<EmailSummary>>(
                await tools.outlook_search_emails(folder: "School/Kids/Sophie"))
            .Should().ContainSingle(message => message.Id == "m4");

        await tools.outlook_move_email("m1", "School/Kids/Sophie");
        ToolResults.Ok<List<EmailSummary>>(
                await tools.outlook_search_emails(folder: "School/Kids/Sophie"))
            .Should().Contain(message => message.Id == "m1");
    }

    [Fact]
    public async Task Tool_errors_carry_codes_and_hints_as_isError()
    {
        var tools = Create(new FakeGraphMailService());

        ToolResults.Fail(await tools.outlook_read_email("nope")).Should().Contain("[message-not-found]");
        ToolResults.Fail(await tools.outlook_read_email("nope")).Should().Contain("outlook_search_emails");

        ToolResults.Fail(await tools.outlook_move_email("m1", "Nirwana")).Should().Contain("[folder-not-found]");
        ToolResults.Fail(await tools.outlook_move_email("m1", "Nirwana")).Should().Contain("outlook_list_folders");

        ToolResults.Fail(await tools.outlook_read_email("  ")).Should().Contain("[invalid-request]");
        ToolResults.Fail(await tools.outlook_create_draft([], "s", "b")).Should().Contain("[invalid-request]");
    }

    [Fact]
    public async Task Unexpected_backend_failures_are_mapped_not_leaked()
    {
        IGraphMailService failing = Substitute.For<IGraphMailService>();
        failing.GetAsync(Arg.Any<string>(), Arg.Any<CancellationToken>())
            .Returns<Task<EmailDetail>>(_ => throw new HttpRequestException("no route"));
        var tools = Create(failing);

        var text = ToolResults.Fail(await tools.outlook_read_email("m1"));
        text.Should().Contain("[service-unavailable]");
        text.Should().Contain("Next:");
    }

    [Fact]
    public void ToolResult_fail_shape_is_agent_readable()
    {
        var result = ToolResult.Fail(
            MicrosoftMcp.Common.GraphServiceException.Throttled(null));

        Assert.True(result.IsError is true);
        ToolResult.ReadText(result).Should().Contain("[throttled]");
    }
}
