using System.Text.Json;
using AwesomeAssertions;
using Microsoft.Extensions.Logging.Abstractions;
using ModelContextProtocol.Protocol;
using NSubstitute;

namespace MicrosoftMcp.Teams.Tests;

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

public sealed class TeamsToolsTests
{
    private static TeamsTools Create(IGraphTeamsService teams) =>
        new(teams, NullLogger<TeamsTools>.Instance);

    [Fact]
    public async Task Browse_teams_channels_threads_with_fake()
    {
        var tools = Create(new FakeGraphTeamsService());

        var teams = ToolResults.Ok<List<TeamInfo>>(await tools.teams_list_teams());
        teams.Should().ContainSingle(t => t.Id == "t-eng");

        var channels = ToolResults.Ok<List<ChannelInfo>>(await tools.teams_list_channels("t-eng"));
        channels.Should().HaveCount(2);

        var messages = ToolResults.Ok<List<MessageSummary>>(
            await tools.teams_list_channel_messages("t-eng", "c-general"));
        messages.Should().ContainSingle(m => m.Id == "m-hello");

        var replies = ToolResults.Ok<List<MessageSummary>>(
            await tools.teams_list_message_replies("t-eng", "c-general", "m-hello"));
        replies.Should().ContainSingle(m => m.Id == "m-r1");

        var detail = ToolResults.Ok<MessageDetail>(
            await tools.teams_read_channel_message("t-eng", "c-general", "m-hello"));
        detail.From.Should().Be("Alice");
    }

    [Fact]
    public async Task Browse_chats_with_fake()
    {
        var tools = Create(new FakeGraphTeamsService());

        var chats = ToolResults.Ok<List<ChatInfo>>(await tools.teams_list_chats());
        chats.Should().ContainSingle(c => c.Id == "chat-1");

        var messages = ToolResults.Ok<List<MessageSummary>>(
            await tools.teams_list_chat_messages("chat-1"));
        messages.Should().ContainSingle(m => m.Id == "cm-1");

        var detail = ToolResults.Ok<MessageDetail>(
            await tools.teams_read_chat_message("chat-1", "cm-1"));
        detail.Id.Should().Be("cm-1");
    }

    [Fact]
    public async Task Tool_errors_carry_codes_and_hints()
    {
        var tools = Create(new FakeGraphTeamsService());

        ToolResults.Fail(await tools.teams_list_channels("no-team")).Should().Contain("[team-not-found]");
        ToolResults.Fail(await tools.teams_list_channel_messages("t-eng", "no-channel"))
            .Should().Contain("[channel-not-found]");
        ToolResults.Fail(await tools.teams_read_channel_message("t-eng", "c-general", "nope"))
            .Should().Contain("[channel-message-not-found]");
        ToolResults.Fail(await tools.teams_list_chat_messages("no-chat")).Should().Contain("[chat-not-found]");
        ToolResults.Fail(await tools.teams_read_chat_message("chat-1", "  ")).Should().Contain("[invalid-request]");
    }

    [Fact]
    public async Task Unexpected_backend_failures_are_mapped_not_leaked()
    {
        IGraphTeamsService failing = Substitute.For<IGraphTeamsService>();
        failing.ListTeamsAsync(Arg.Any<CancellationToken>())
            .Returns<Task<IReadOnlyList<TeamInfo>>>(_ => throw new HttpRequestException("no route"));
        var tools = Create(failing);

        var text = ToolResults.Fail(await tools.teams_list_teams());
        text.Should().Contain("[service-unavailable]");
        text.Should().Contain("Next:");
    }

    [Fact]
    public async Task Delegation_wiring_with_substitute()
    {
        IGraphTeamsService teams = Substitute.For<IGraphTeamsService>();
        teams.ListTeamsAsync(Arg.Any<CancellationToken>()).Returns(
            [new TeamInfo("t1", "Team", null, "private", false)]);
        var tools = Create(teams);

        ToolResults.Ok<List<TeamInfo>>(await tools.teams_list_teams())
            .Should().ContainSingle(t => t.Id == "t1");
    }
}
