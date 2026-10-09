using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Serialization;
using MicrosoftMcp.Common;
using System.Text.Json;

namespace MicrosoftMcp.Teams;

/// <summary>Read-only Teams access for the signed-in user.</summary>
public sealed class GraphTeamsService : IGraphTeamsService
{
    private readonly GraphServiceClient _client;
    private readonly string _userIdOrUpn;

    public GraphTeamsService(
        GraphServiceClient client,
        IOptions<GraphAuthOptions> options)
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
        var page = await _client.Me.JoinedTeams.GetAsync(cancellationToken: ct).ConfigureAwait(false);
        return [.. (page?.Value ?? []).Select(TeamsMapper.MapTeam)];
    }

    public async Task<IReadOnlyList<ChannelInfo>> ListChannelsAsync(string teamId, CancellationToken ct = default)
    {
        RequireId(teamId, "teamId", "call teams_list_teams to get valid team ids");
        string normalizedTeamId = teamId.Trim();
        var page = await _client.Teams[normalizedTeamId].Channels.GetAsync(c =>
        {
            c.QueryParameters.Select = ["id", "displayName", "description", "membershipType"];
        }, ct).ConfigureAwait(false);
        var channels = new List<Channel>();
        while (page is not null)
        {
            channels.AddRange(page.Value ?? []);
            if (string.IsNullOrWhiteSpace(page.OdataNextLink))
            {
                break;
            }

            page = await _client.Teams[normalizedTeamId].Channels
                .WithUrl(page.OdataNextLink)
                .GetAsync(cancellationToken: ct)
                .ConfigureAwait(false);
        }

        return [.. channels.Select(TeamsMapper.MapChannel)];
    }

    public async Task<ChannelFilesFolderInfo> GetChannelFilesFolderAsync(
        string teamId, string channelId, CancellationToken ct = default)
    {
        RequireId(teamId, "teamId", "call teams_list_teams to get valid team ids");
        RequireId(channelId, "channelId", "call teams_list_channels for the team to get valid channel ids");
        string tid = teamId.Trim();
        string cid = channelId.Trim();
        // NOTE: the v1.0 filesFolder endpoint documents no OData query
        // parameters, so no $select is sent (it would fail with HTTP 400).
        var folder = await _client.Teams[tid].Channels[cid].FilesFolder
            .GetAsync(cancellationToken: ct).ConfigureAwait(false);
        if (folder?.Id is null)
        {
            throw GraphServiceException.ChannelFilesFolderNotFound(channelId, "teams_get_channel_files_folder");
        }

        if (string.IsNullOrWhiteSpace(folder.ParentReference?.DriveId))
        {
            throw GraphServiceException.InvalidRequest(
                $"Files folder for channel '{channelId}' has no drive id.",
                "the channel storage may not be provisioned yet; open the Files tab in Teams once, then retry");
        }

        return new ChannelFilesFolderInfo(
            tid,
            cid,
            folder.ParentReference?.DriveId,
            folder.Id,
            folder.Name,
            folder.ParentReference?.SiteId,
            folder.WebUrl);
    }

    public async Task<IReadOnlyList<MessageSummary>> ListChannelMessagesAsync(
        string teamId, string channelId, int top = 25, CancellationToken ct = default)
    {
        RequireId(teamId, "teamId", "call teams_list_teams to get valid team ids");
        RequireId(channelId, "channelId", "call teams_list_channels for the team to get valid channel ids");
        int take = Math.Clamp(top, 1, 50);
        var page = await _client.Teams[teamId.Trim()].Channels[channelId.Trim()].Messages.GetAsync(c =>
        {
            c.QueryParameters.Top = take;
        }, ct).ConfigureAwait(false);
        return [.. (page?.Value ?? []).Select(TeamsMapper.MapSummary)];
    }

    public async Task<IReadOnlyList<MessageSummary>> ListMessageRepliesAsync(
        string teamId, string channelId, string messageId, int top = 25, CancellationToken ct = default)
    {
        RequireId(teamId, "teamId", "call teams_list_teams to get valid team ids");
        RequireId(channelId, "channelId", "call teams_list_channels for the team to get valid channel ids");
        RequireId(messageId, "messageId", "call teams_list_channel_messages to get valid message ids");
        int take = Math.Clamp(top, 1, 50);
        var page = await _client.Teams[teamId.Trim()].Channels[channelId.Trim()].Messages[messageId.Trim()]
            .Replies.GetAsync(c =>
            {
                c.QueryParameters.Top = take;
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
            .GetAsync(cancellationToken: ct).ConfigureAwait(false);
        return msg is null
            ? throw GraphServiceException.TeamsMessageNotFound(messageId, "teams_read_channel_message")
            : TeamsMapper.MapDetail(msg);
    }

    public async Task<IReadOnlyList<ChatInfo>> ListChatsAsync(int top = 25, CancellationToken ct = default)
    {
        int take = Math.Clamp(top, 1, 50);
        var page = await _client.Me.Chats.GetAsync(c =>
        {
            c.QueryParameters.Top = take;
        }, ct).ConfigureAwait(false);
        return [.. (page?.Value ?? []).Select(TeamsMapper.MapChat)];
    }

    public async Task<IReadOnlyList<MessageSummary>> ListChatMessagesAsync(
        string chatId, int top = 25, CancellationToken ct = default)
    {
        RequireId(chatId, "chatId", "call teams_list_chats to get valid chat ids");
        int take = Math.Clamp(top, 1, 50);
        var page = await _client.Me.Chats[chatId.Trim()].Messages.GetAsync(c =>
        {
            c.QueryParameters.Top = take;
        }, ct).ConfigureAwait(false);
        return [.. (page?.Value ?? []).Select(TeamsMapper.MapSummary)];
    }

    public async Task<MessageDetail> ReadChatMessageAsync(
        string chatId, string messageId, CancellationToken ct = default)
    {
        RequireId(chatId, "chatId", "call teams_list_chats to get valid chat ids");
        RequireId(messageId, "messageId", "call teams_list_chat_messages to get valid message ids");
        var msg = await _client.Me.Chats[chatId.Trim()].Messages[messageId.Trim()]
            .GetAsync(cancellationToken: ct).ConfigureAwait(false);
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
        string userId = await ResolveUserIdAsync(ct).ConfigureAwait(false);
        string nextLink = $"/copilot/users/{Uri.EscapeDataString(userId)}/onlineMeetings/"
            + $"{Uri.EscapeDataString(meetingId.Trim())}/aiInsights";
        var insights = new List<MeetingInsightInfo>();
        while (!string.IsNullOrWhiteSpace(nextLink) && insights.Count < take)
        {
            using JsonDocument document = await SendJsonAsync(nextLink, ct).ConfigureAwait(false);
            foreach (JsonElement insight in ReadValue(document.RootElement))
            {
                insights.Add(TeamsMapper.MapInsightSummary(insight));
                if (insights.Count == take)
                {
                    break;
                }
            }

            nextLink = insights.Count == take
                ? string.Empty
                : NextLink(document.RootElement) ?? string.Empty;
        }

        return insights;
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

    private async Task<JsonDocument> GetInsightAsync(
        string meetingId, string insightId, string userId, CancellationToken ct)
    {
        return await SendJsonAsync(
            $"/copilot/users/{Uri.EscapeDataString(userId)}/onlineMeetings/"
                + $"{Uri.EscapeDataString(meetingId)}/aiInsights/{Uri.EscapeDataString(insightId)}",
            ct).ConfigureAwait(false);
    }

    private async Task<JsonDocument> SendJsonAsync(string path, CancellationToken ct)
    {
        var adapter = _client.RequestAdapter
            ?? throw GraphServiceException.GraphError(0, null, "Graph request adapter is unavailable.");
        string baseUrl = adapter.BaseUrl?.TrimEnd('/')
            ?? throw GraphServiceException.GraphError(0, null, "Graph request adapter has no base URL.");
        string requestPath = path;
        if (Uri.TryCreate(path, UriKind.Absolute, out var absoluteUri))
        {
            Uri baseUri = new(baseUrl + "/");
            string basePath = baseUri.AbsolutePath.TrimEnd('/');
            requestPath = absoluteUri.AbsolutePath.StartsWith(basePath, StringComparison.OrdinalIgnoreCase)
                ? absoluteUri.AbsolutePath[basePath.Length..] + absoluteUri.Query
                : absoluteUri.PathAndQuery;
        }

        var request = new RequestInformation
        {
            HttpMethod = Method.GET,
            UrlTemplate = "{+baseurl}" + requestPath,
            PathParameters = new Dictionary<string, object>
            {
                ["baseurl"] = baseUrl
            }
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

    private static string? NextLink(JsonElement root) =>
        root.TryGetProperty("@odata.nextLink", out var nextLink)
            && nextLink.ValueKind == JsonValueKind.String
            ? nextLink.GetString()
            : null;

    private static void RequireId(string value, string what, string hint)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw GraphServiceException.InvalidRequest($"{what} must not be empty.", hint);
        }
    }
}
