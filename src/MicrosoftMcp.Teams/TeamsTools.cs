using System.ComponentModel;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using MicrosoftMcp.Common;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace MicrosoftMcp.Teams;

public static class TeamsServiceRegistration
{
    public static IServiceCollection AddTeams(this IServiceCollection services)
    {
        services.AddSingleton<IGraphTeamsService, GraphTeamsService>();
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
        catch (MailServiceException ex)
        {
            log.LogWarning(ex, "Tool {Operation} failed with {Code}", operation, ex.Code);
            return ToolResult.Fail(ex);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            var mapped = GraphErrorMapper.ToMailServiceException(ex, operation, resource);
            log.LogError(ex, "Tool {Operation} failed unexpectedly ({Code})", operation, mapped.Code);
            return ToolResult.Fail(mapped);
        }
    }

    [McpServerTool, Description("List joined teams (id, name). Needed to pick a team id. Read-only.")]
    public Task<CallToolResult> list_teams(CancellationToken ct = default) =>
        InvokeAsync("list_teams", () => teams.ListTeamsAsync(ct), "team");

    [McpServerTool, Description("List channels of a team. Read-only.")]
    public Task<CallToolResult> list_channels(
        [Description("Team id from list_teams")] string teamId,
        CancellationToken ct = default) =>
        InvokeAsync("list_channels", () => teams.ListChannelsAsync(teamId, ct), "channel");

    [McpServerTool, Description("List messages of a channel (newest first). Read-only, no send.")]
    public Task<CallToolResult> list_channel_messages(
        [Description("Team id")] string teamId,
        [Description("Channel id")] string channelId,
        [Description("Max messages 1-100")] int top = 25,
        CancellationToken ct = default) =>
        InvokeAsync("list_channel_messages", () => teams.ListChannelMessagesAsync(teamId, channelId, top, ct));

    [McpServerTool, Description("List replies to a channel message (thread). Read-only.")]
    public Task<CallToolResult> list_message_replies(
        [Description("Team id")] string teamId,
        [Description("Channel id")] string channelId,
        [Description("Parent message id")] string messageId,
        [Description("Max replies 1-100")] int top = 25,
        CancellationToken ct = default) =>
        InvokeAsync("list_message_replies", () => teams.ListMessageRepliesAsync(teamId, channelId, messageId, top, ct));

    [McpServerTool, Description("Read a full channel message (content, reactions, mentions). Read-only.")]
    public Task<CallToolResult> read_channel_message(
        [Description("Team id")] string teamId,
        [Description("Channel id")] string channelId,
        [Description("Message id")] string messageId,
        CancellationToken ct = default) =>
        InvokeAsync("read_channel_message", () => teams.ReadChannelMessageAsync(teamId, channelId, messageId, ct));

    [McpServerTool, Description("List recent chats (1:1 and group). Read-only.")]
    public Task<CallToolResult> list_chats(
        [Description("Max chats 1-100")] int top = 25,
        CancellationToken ct = default) =>
        InvokeAsync("list_chats", () => teams.ListChatsAsync(top, ct), "chat");

    [McpServerTool, Description("List messages of a chat. Read-only, no send.")]
    public Task<CallToolResult> list_chat_messages(
        [Description("Chat id from list_chats")] string chatId,
        [Description("Max messages 1-100")] int top = 25,
        CancellationToken ct = default) =>
        InvokeAsync("list_chat_messages", () => teams.ListChatMessagesAsync(chatId, top, ct));

    [McpServerTool, Description("List replies to a chat message (thread). Read-only.")]
    public Task<CallToolResult> list_chat_replies(
        [Description("Chat id")] string chatId,
        [Description("Parent message id")] string messageId,
        [Description("Max replies 1-100")] int top = 25,
        CancellationToken ct = default) =>
        InvokeAsync("list_chat_replies", () => teams.ListChatRepliesAsync(chatId, messageId, top, ct));

    [McpServerTool, Description("Read a full chat message (content, reactions, mentions). Read-only.")]
    public Task<CallToolResult> read_chat_message(
        [Description("Chat id")] string chatId,
        [Description("Message id")] string messageId,
        CancellationToken ct = default) =>
        InvokeAsync("read_chat_message", () => teams.ReadChatMessageAsync(chatId, messageId, ct));
}
