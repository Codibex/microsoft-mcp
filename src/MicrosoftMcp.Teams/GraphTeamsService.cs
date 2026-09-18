using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Serialization;
using MicrosoftMcp.Common;
using System.Text.Json;

namespace MicrosoftMcp.Teams;

/// <summary>Read-only Teams access for the signed-in user. No send and no
/// writes exist on this service by design.</summary>
public sealed class GraphTeamsService : IGraphTeamsService
{
    private readonly GraphServiceClient _client;
    private readonly string _userIdOrUpn;

    public GraphTeamsService(GraphServiceClient client, IOptions<GraphAuthOptions> options)
    {
        // Teamwork APIs via /me/* require a signed-in user. App-only would
        // need a different permission/path scheme (out of scope).
        if (options.Value.AuthMode == AuthMode.AppOnly)
        {
            throw GraphServiceException.AuthMisconfigured(
                "The Teams host requires delegated auth. Set Graph:AuthMode to Delegated.");
        }

        _client = client;
        _userIdOrUpn = options.Value.UserIdOrUpn;
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
        RequireId(teamId, "teamId", "call teams_list_teams to get valid team ids");
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
        RequireId(teamId, "teamId", "call teams_list_teams to get valid team ids");
        RequireId(channelId, "channelId", "call teams_list_channels for the team to get valid channel ids");
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
        RequireId(teamId, "teamId", "call teams_list_teams to get valid team ids");
        RequireId(channelId, "channelId", "call teams_list_channels for the team to get valid channel ids");
        RequireId(messageId, "messageId", "call teams_list_channel_messages to get valid message ids");
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
        RequireId(teamId, "teamId", "call teams_list_teams to get valid team ids");
        RequireId(channelId, "channelId", "call teams_list_channels for the team to get valid channel ids");
        RequireId(messageId, "messageId", "call teams_list_channel_messages to get valid message ids");
        var msg = await _client.Teams[teamId.Trim()].Channels[channelId.Trim()].Messages[messageId.Trim()]
            .GetAsync(c => c.QueryParameters.Select = TeamsMapper.Select, ct).ConfigureAwait(false);
        return msg is null
            ? throw GraphServiceException.TeamsMessageNotFound(messageId, "teams_read_channel_message")
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
        RequireId(chatId, "chatId", "call teams_list_chats to get valid chat ids");
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
        RequireId(chatId, "chatId", "call teams_list_chats to get valid chat ids");
        RequireId(messageId, "messageId", "call teams_list_chat_messages to get valid message ids");
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
        RequireId(chatId, "chatId", "call teams_list_chats to get valid chat ids");
        RequireId(messageId, "messageId", "call teams_list_chat_messages to get valid message ids");
        var msg = await _client.Me.Chats[chatId.Trim()].Messages[messageId.Trim()]
            .GetAsync(c => c.QueryParameters.Select = TeamsMapper.Select, ct).ConfigureAwait(false);
        return msg is null
            ? throw GraphServiceException.TeamsMessageNotFound(messageId, "teams_read_chat_message")
            : TeamsMapper.MapDetail(msg);
    }

    public async Task<IReadOnlyList<MeetingTranscriptInfo>> ListMeetingTranscriptsAsync(
        string meetingId, int top = 25, CancellationToken ct = default)
    {
        RequireId(meetingId, "meetingId", "use the online meeting id from Microsoft Graph");
        int take = Math.Clamp(top, 1, 100);
        var page = await _client.Me.OnlineMeetings[meetingId.Trim()].Transcripts.GetAsync(c =>
        {
            c.QueryParameters.Top = take;
            c.QueryParameters.Select = [
                "id", "meetingId", "callId", "createdDateTime", "endDateTime", "contentCorrelationId"
            ];
        }, ct).ConfigureAwait(false);
        return [.. (page?.Value ?? []).Select(TeamsMapper.MapTranscript)];
    }

    public async Task<MeetingTranscriptDetail> ReadMeetingTranscriptAsync(
        string meetingId, string transcriptId, CancellationToken ct = default)
    {
        RequireId(meetingId, "meetingId", "use the online meeting id from Microsoft Graph");
        RequireId(transcriptId, "transcriptId", "call teams_list_meeting_transcripts to get valid transcript ids");
        var transcript = await _client.Me.OnlineMeetings[meetingId.Trim()].Transcripts[transcriptId.Trim()]
            .GetAsync(c => c.QueryParameters.Select = [
                "id", "meetingId", "callId", "createdDateTime", "endDateTime", "contentCorrelationId"
            ], ct).ConfigureAwait(false);
        if (transcript is null)
        {
            throw GraphServiceException.MeetingTranscriptNotFound(transcriptId, "teams_read_meeting_transcript");
        }

        using Stream? stream = await _client.Me.OnlineMeetings[meetingId.Trim()].Transcripts[transcriptId.Trim()]
            .Content.GetAsync(c => c.Headers.Add("Accept", "text/vtt"), ct).ConfigureAwait(false);
        if (stream is null)
        {
            throw GraphServiceException.GraphError(
                200, null, "Graph returned no meeting transcript content.");
        }

        string content = await new StreamReader(stream).ReadToEndAsync(ct).ConfigureAwait(false);
        if (string.IsNullOrWhiteSpace(content))
        {
            throw GraphServiceException.GraphError(
                200, null, "Graph returned an empty meeting transcript response.");
        }

        return TeamsMapper.MapTranscriptDetail(transcript, content);
    }

    public async Task<IReadOnlyList<MeetingInsightInfo>> ListMeetingInsightsAsync(
        string meetingId, int top = 25, CancellationToken ct = default)
    {
        RequireId(meetingId, "meetingId", "use the online meeting id from Microsoft Graph");
        int take = Math.Clamp(top, 1, 100);
        using JsonDocument document = await GetInsightsAsync(meetingId.Trim(), take, ct).ConfigureAwait(false);
        return [.. ReadValue(document.RootElement).Select(TeamsMapper.MapInsightSummary)];
    }

    public async Task<MeetingInsightDetail> ReadMeetingInsightAsync(
        string meetingId, string insightId, CancellationToken ct = default)
    {
        RequireId(meetingId, "meetingId", "use the online meeting id from Microsoft Graph");
        RequireId(insightId, "insightId", "call teams_list_meeting_insights to get valid insight ids");
        string userId = await ResolveUserIdAsync(ct).ConfigureAwait(false);
        using JsonDocument document = await GetInsightAsync(
            meetingId.Trim(), insightId.Trim(), userId, ct).ConfigureAwait(false);
        JsonElement insight = document.RootElement;
        if (insight.ValueKind != JsonValueKind.Object
            || !insight.TryGetProperty("id", out var id)
            || !string.Equals(id.GetString(), insightId.Trim(), StringComparison.Ordinal))
        {
            throw GraphServiceException.MeetingInsightNotFound(insightId, "teams_read_meeting_insight");
        }

        return TeamsMapper.MapInsightDetail(insight);
    }

    private async Task<JsonDocument> GetInsightsAsync(string meetingId, int top, CancellationToken ct)
    {
        string userId = await ResolveUserIdAsync(ct).ConfigureAwait(false);
        return await GetInsightsAsync(meetingId, $"$top={top}", userId, ct).ConfigureAwait(false);
    }

    private async Task<JsonDocument> GetInsightAsync(
        string meetingId, string insightId, string userId, CancellationToken ct)
    {
        return await SendJsonAsync(
            $"/copilot/users/{Uri.EscapeDataString(userId)}/onlineMeetings/"
                + $"{Uri.EscapeDataString(meetingId)}/aiInsights/{Uri.EscapeDataString(insightId)}",
            null,
            ct).ConfigureAwait(false);
    }

    private async Task<JsonDocument> GetInsightsAsync(
        string meetingId, string query, string userId, CancellationToken ct)
    {
        return await SendJsonAsync(
            $"/copilot/users/{Uri.EscapeDataString(userId)}/onlineMeetings/"
                + $"{Uri.EscapeDataString(meetingId)}/aiInsights",
            query,
            ct).ConfigureAwait(false);
    }

    private async Task<JsonDocument> SendJsonAsync(string path, string? query, CancellationToken ct)
    {
        var adapter = _client.RequestAdapter
            ?? throw GraphServiceException.GraphError(0, null, "Graph request adapter is unavailable.");
        string baseUrl = adapter.BaseUrl?.TrimEnd('/')
            ?? throw GraphServiceException.GraphError(0, null, "Graph request adapter has no base URL.");
        string uri = baseUrl + path;
        if (!string.IsNullOrWhiteSpace(query))
        {
            uri += $"?{query}";
        }

        var request = new RequestInformation
        {
            HttpMethod = Method.GET,
            URI = new Uri(uri)
        };
        request.Headers.Add("Accept", "application/json");
        using Stream? stream = await adapter.SendPrimitiveAsync<Stream>(
            request,
            new Dictionary<string, ParsableFactory<IParsable>>(),
            ct).ConfigureAwait(false);
        if (stream is null)
        {
            throw GraphServiceException.GraphError(200, null, "Graph returned an empty AI insights response.");
        }

        return await JsonDocument.ParseAsync(stream, cancellationToken: ct).ConfigureAwait(false);
    }

    private async Task<string> ResolveUserIdAsync(CancellationToken ct)
    {
        if (!string.IsNullOrWhiteSpace(_userIdOrUpn)
            && !string.Equals(_userIdOrUpn, "me", StringComparison.OrdinalIgnoreCase))
        {
            return _userIdOrUpn.Trim();
        }

        var user = await _client.Me.GetAsync(c => c.QueryParameters.Select = ["id"], ct)
            .ConfigureAwait(false);
        return !string.IsNullOrWhiteSpace(user?.Id)
            ? user.Id
            : throw GraphServiceException.AuthMisconfigured(
                "Microsoft Graph did not return the signed-in user's id for meeting insights.");
    }

    private static IEnumerable<JsonElement> ReadValue(JsonElement root)
    {
        if (root.ValueKind == JsonValueKind.Array)
        {
            return root.EnumerateArray();
        }

        return root.TryGetProperty("value", out var value) && value.ValueKind == JsonValueKind.Array
            ? value.EnumerateArray()
            : [];
    }

    private static void RequireId(string value, string what, string hint)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw GraphServiceException.InvalidRequest($"{what} must not be empty.", hint);
        }
    }
}
