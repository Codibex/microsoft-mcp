using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using MicrosoftMcp.Common;

namespace MicrosoftMcp.Teams;

/// <summary>Read-only Teams access for the signed-in user. No send and no
/// writes exist on this service by design.</summary>
public sealed class GraphTeamsService : IGraphTeamsService
{
    private readonly GraphServiceClient _client;

    public GraphTeamsService(GraphServiceClient client, IOptions<GraphAuthOptions> options)
    {
        // Teamwork APIs via /me/* require a signed-in user. App-only would
        // need a different permission/path scheme (out of scope).
        if (options.Value.AuthMode == AuthMode.AppOnly)
        {
            throw MailServiceException.AuthMisconfigured(
                "The Teams host requires delegated auth. Set Graph:AuthMode to Delegated.");
        }

        _client = client;
    }

    public async Task<IReadOnlyList<TeamInfo>> ListTeamsAsync(CancellationToken ct = default)
    {
        var page = await _client.Me.JoinedTeams.GetAsync(c =>
        {
            c.QueryParameters.Top = 100;
            c.QueryParameters.Select = ["id", "displayName", "description", "visibility", "isArchived"];
        }, ct).ConfigureAwait(false);
        return [.. (page?.Value ?? []).Select(TeamsMapper.MapTeam)];
    }

    public async Task<IReadOnlyList<ChannelInfo>> ListChannelsAsync(string teamId, CancellationToken ct = default)
    {
        RequireId(teamId, "teamId", "call list_teams to get valid team ids");
        var page = await _client.Teams[teamId.Trim()].Channels.GetAsync(c =>
        {
            c.QueryParameters.Top = 100;
            c.QueryParameters.Select = ["id", "displayName", "description", "membershipType"];
        }, ct).ConfigureAwait(false);
        return [.. (page?.Value ?? []).Select(TeamsMapper.MapChannel)];
    }

    public async Task<IReadOnlyList<MessageSummary>> ListChannelMessagesAsync(
        string teamId, string channelId, int top = 25, CancellationToken ct = default)
    {
        RequireId(teamId, "teamId", "call list_teams to get valid team ids");
        RequireId(channelId, "channelId", "call list_channels for the team to get valid channel ids");
        int take = Math.Clamp(top, 1, 100);
        var page = await _client.Teams[teamId.Trim()].Channels[channelId.Trim()].Messages.GetAsync(c =>
        {
            c.QueryParameters.Top = take;
            c.QueryParameters.Select = TeamsMapper.Select;
        }, ct).ConfigureAwait(false);
        return [.. (page?.Value ?? []).Select(TeamsMapper.MapSummary)];
    }

    public async Task<IReadOnlyList<MessageSummary>> ListMessageRepliesAsync(
        string teamId, string channelId, string messageId, int top = 25, CancellationToken ct = default)
    {
        RequireId(teamId, "teamId", "call list_teams to get valid team ids");
        RequireId(channelId, "channelId", "call list_channels for the team to get valid channel ids");
        RequireId(messageId, "messageId", "call list_channel_messages to get valid message ids");
        int take = Math.Clamp(top, 1, 100);
        var page = await _client.Teams[teamId.Trim()].Channels[channelId.Trim()].Messages[messageId.Trim()]
            .Replies.GetAsync(c =>
            {
                c.QueryParameters.Top = take;
                c.QueryParameters.Select = TeamsMapper.Select;
            }, ct).ConfigureAwait(false);
        return [.. (page?.Value ?? []).Select(TeamsMapper.MapSummary)];
    }

    public async Task<MessageDetail> ReadChannelMessageAsync(
        string teamId, string channelId, string messageId, CancellationToken ct = default)
    {
        RequireId(teamId, "teamId", "call list_teams to get valid team ids");
        RequireId(channelId, "channelId", "call list_channels for the team to get valid channel ids");
        RequireId(messageId, "messageId", "call list_channel_messages to get valid message ids");
        var msg = await _client.Teams[teamId.Trim()].Channels[channelId.Trim()].Messages[messageId.Trim()]
            .GetAsync(c => c.QueryParameters.Select = TeamsMapper.Select, ct).ConfigureAwait(false);
        return msg is null
            ? throw MailServiceException.TeamsMessageNotFound(messageId, "read_channel_message")
            : TeamsMapper.MapDetail(msg);
    }

    public async Task<IReadOnlyList<ChatInfo>> ListChatsAsync(int top = 25, CancellationToken ct = default)
    {
        int take = Math.Clamp(top, 1, 100);
        var page = await _client.Me.Chats.GetAsync(c =>
        {
            c.QueryParameters.Top = take;
            c.QueryParameters.Select = ["id", "topic", "chatType", "lastUpdatedDateTime"];
        }, ct).ConfigureAwait(false);
        return [.. (page?.Value ?? []).Select(TeamsMapper.MapChat)];
    }

    public async Task<IReadOnlyList<MessageSummary>> ListChatMessagesAsync(
        string chatId, int top = 25, CancellationToken ct = default)
    {
        RequireId(chatId, "chatId", "call list_chats to get valid chat ids");
        int take = Math.Clamp(top, 1, 100);
        var page = await _client.Me.Chats[chatId.Trim()].Messages.GetAsync(c =>
        {
            c.QueryParameters.Top = take;
            c.QueryParameters.Select = TeamsMapper.Select;
        }, ct).ConfigureAwait(false);
        return [.. (page?.Value ?? []).Select(TeamsMapper.MapSummary)];
    }

    public async Task<IReadOnlyList<MessageSummary>> ListChatRepliesAsync(
        string chatId, string messageId, int top = 25, CancellationToken ct = default)
    {
        RequireId(chatId, "chatId", "call list_chats to get valid chat ids");
        RequireId(messageId, "messageId", "call list_chat_messages to get valid message ids");
        int take = Math.Clamp(top, 1, 100);
        var page = await _client.Me.Chats[chatId.Trim()].Messages[messageId.Trim()].Replies.GetAsync(c =>
        {
            c.QueryParameters.Top = take;
            c.QueryParameters.Select = TeamsMapper.Select;
        }, ct).ConfigureAwait(false);
        return [.. (page?.Value ?? []).Select(TeamsMapper.MapSummary)];
    }

    public async Task<MessageDetail> ReadChatMessageAsync(
        string chatId, string messageId, CancellationToken ct = default)
    {
        RequireId(chatId, "chatId", "call list_chats to get valid chat ids");
        RequireId(messageId, "messageId", "call list_chat_messages to get valid message ids");
        var msg = await _client.Me.Chats[chatId.Trim()].Messages[messageId.Trim()]
            .GetAsync(c => c.QueryParameters.Select = TeamsMapper.Select, ct).ConfigureAwait(false);
        return msg is null
            ? throw MailServiceException.TeamsMessageNotFound(messageId, "read_chat_message")
            : TeamsMapper.MapDetail(msg);
    }

    private static void RequireId(string value, string what, string hint)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw MailServiceException.InvalidRequest($"{what} must not be empty.", hint);
        }
    }
}
