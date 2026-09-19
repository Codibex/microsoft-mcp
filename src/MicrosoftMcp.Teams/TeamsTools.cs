using System.ComponentModel;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using MicrosoftMcp.Common;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace MicrosoftMcp.Teams;

public static class TeamsServiceRegistration
{
    public static IServiceCollection AddTeams(
        this IServiceCollection services, TeamsPolicyOptions? policy = null)
    {
        services.AddSingleton<IGraphTeamsService, GraphTeamsService>();
        services.AddSingleton<IOptions<TeamsPolicyOptions>>(
            Options.Create(policy ?? new TeamsPolicyOptions()));
        return services;
    }
}

[McpServerToolType]
public sealed class TeamsTools(IGraphTeamsService teams, ILogger<TeamsTools> log)
{
    private async Task<CallToolResult> InvokeAsync<T>(
        string operation, Func<Task<T>> call, string resource = "teams-message")
    {
        try
        {
            return ToolResult.Ok(await call().ConfigureAwait(false));
        }
        catch (GraphServiceException ex)
        {
            log.LogWarning(ex, "Tool {Operation} failed with {Code}", operation, ex.Code);
            return ToolResult.Fail(ex);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            var mapped = GraphErrorMapper.ToGraphServiceException(ex, operation, resource);
            log.LogError(ex, "Tool {Operation} failed unexpectedly ({Code})", operation, mapped.Code);
            return ToolResult.Fail(mapped);
        }
    }

    [McpServerTool, Description("List joined teams (id, name). Needed to pick a team id. Read-only.")]
    public Task<CallToolResult> teams_list_teams(CancellationToken ct = default) =>
        InvokeAsync("teams_list_teams", () => teams.ListTeamsAsync(ct), "team");

    [McpServerTool, Description("List channels of a team. Read-only.")]
    public Task<CallToolResult> teams_list_channels(
        [Description("Team id from teams_list_teams")] string teamId,
        CancellationToken ct = default) =>
        InvokeAsync("teams_list_channels", () => teams.ListChannelsAsync(teamId, ct), "channel");

    [McpServerTool, Description("List messages of a channel (newest first).")]
    public Task<CallToolResult> teams_list_channel_messages(
        [Description("Team id")] string teamId,
        [Description("Channel id")] string channelId,
        [Description("Max messages 1-50")] int top = 25,
        CancellationToken ct = default) =>
        InvokeAsync("teams_list_channel_messages", () => teams.ListChannelMessagesAsync(teamId, channelId, top, ct));

    [McpServerTool, Description("List replies to a channel message (thread). Read-only.")]
    public Task<CallToolResult> teams_list_message_replies(
        [Description("Team id")] string teamId,
        [Description("Channel id")] string channelId,
        [Description("Parent message id")] string messageId,
        [Description("Max replies 1-50")] int top = 25,
        CancellationToken ct = default) =>
        InvokeAsync("teams_list_message_replies", () => teams.ListMessageRepliesAsync(teamId, channelId, messageId, top, ct));

    [McpServerTool, Description("Read a full channel message (content, reactions, mentions). Read-only.")]
    public Task<CallToolResult> teams_read_channel_message(
        [Description("Team id")] string teamId,
        [Description("Channel id")] string channelId,
        [Description("Message id")] string messageId,
        CancellationToken ct = default) =>
        InvokeAsync("teams_read_channel_message", () => teams.ReadChannelMessageAsync(teamId, channelId, messageId, ct));

    [McpServerTool, Description("Send a message to an existing team channel. The server applies policy and disclosure checks.")]
    public Task<CallToolResult> teams_send_channel_message(
        [Description("Team id from teams_list_teams")] string teamId,
        [Description("Channel id from teams_list_channels")] string channelId,
        [Description("Message body")] string body,
        CancellationToken ct = default) =>
        InvokeAsync(
            "teams_send_channel_message",
            () => teams.SendChannelMessageAsync(teamId, channelId, body, ct));

    [McpServerTool, Description("List recent chats (1:1 and group). Read-only.")]
    public Task<CallToolResult> teams_list_chats(
        [Description("Max chats 1-50")] int top = 25,
        CancellationToken ct = default) =>
        InvokeAsync("teams_list_chats", () => teams.ListChatsAsync(top, ct), "chat");

    [McpServerTool, Description("List messages of a chat.")]
    public Task<CallToolResult> teams_list_chat_messages(
        [Description("Chat id from teams_list_chats")] string chatId,
        [Description("Max messages 1-50")] int top = 25,
        CancellationToken ct = default) =>
        InvokeAsync("teams_list_chat_messages", () => teams.ListChatMessagesAsync(chatId, top, ct));

    [McpServerTool, Description("Read a full chat message (content, reactions, mentions). Read-only.")]
    public Task<CallToolResult> teams_read_chat_message(
        [Description("Chat id")] string chatId,
        [Description("Message id")] string messageId,
        CancellationToken ct = default) =>
        InvokeAsync("teams_read_chat_message", () => teams.ReadChatMessageAsync(chatId, messageId, ct));

    [McpServerTool, Description("Send a message to an existing 1:1 or group chat. The server applies policy and disclosure checks.")]
    public Task<CallToolResult> teams_send_chat_message(
        [Description("Chat id from teams_list_chats")] string chatId,
        [Description("Message body")] string body,
        CancellationToken ct = default) =>
        InvokeAsync(
            "teams_send_chat_message",
            () => teams.SendChatMessageAsync(chatId, body, ct));

    [McpServerTool, Description("List transcripts of a scheduled online meeting. Read-only, available after transcription.")]
    public Task<CallToolResult> teams_list_meeting_transcripts(
        [Description("Online meeting id from Microsoft Graph")] string meetingId,
        [Description("Max transcripts 1-100")] int top = 25,
        CancellationToken ct = default) =>
        InvokeAsync(
            "teams_list_meeting_transcripts",
            () => teams.ListMeetingTranscriptsAsync(meetingId, top, ct),
            "meeting-transcript");

    [McpServerTool, Description("Read the VTT transcript of a scheduled online meeting. Read-only.")]
    public Task<CallToolResult> teams_read_meeting_transcript(
        [Description("Online meeting id")] string meetingId,
        [Description("Transcript id from teams_list_meeting_transcripts")] string transcriptId,
        CancellationToken ct = default) =>
        InvokeAsync(
            "teams_read_meeting_transcript",
            () => teams.ReadMeetingTranscriptAsync(meetingId, transcriptId, ct),
            "meeting-transcript");

    [McpServerTool, Description("List AI-generated insights for a completed online meeting. Requires Microsoft 365 Copilot.")]
    public Task<CallToolResult> teams_list_meeting_insights(
        [Description("Online meeting id from Microsoft Graph")] string meetingId,
        [Description("Max insights 1-100")] int top = 25,
        CancellationToken ct = default) =>
        InvokeAsync(
            "teams_list_meeting_insights",
            () => teams.ListMeetingInsightsAsync(meetingId, top, ct),
            "meeting-insight");

    [McpServerTool, Description("Read AI-generated notes, action items and mentions for a completed online meeting.")]
    public Task<CallToolResult> teams_read_meeting_insight(
        [Description("Online meeting id")] string meetingId,
        [Description("Insight id from teams_list_meeting_insights")] string insightId,
        CancellationToken ct = default) =>
        InvokeAsync(
            "teams_read_meeting_insight",
            () => teams.ReadMeetingInsightAsync(meetingId, insightId, ct),
            "meeting-insight");
}
