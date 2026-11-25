using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Serialization;
using MicrosoftMcp.Common;
using System.Text;
using System.Text.Json;

namespace MicrosoftMcp.Teams;

/// <summary>Read-only Teams access for the signed-in user.</summary>
public sealed class GraphTeamsService : IGraphTeamsService
{
    private readonly GraphServiceClient _client;
    private readonly string _userIdOrUpn;

    /// <summary>Page size sent to the chat listing request; the scan budget
    /// below is expressed in chats and divided by it.</summary>
    private const int ChatPageSize = 50;

    /// <summary>Scan cap so one call cannot page a pathological chat list
    /// forever (same pattern as GraphTodoService.MaxScannedTasks).</summary>
    private const int MaxScannedChats = 500;

    /// <summary>Member cap for teams_list_chat_members (same idea as
    /// MaxScannedChats; group chats stay far below this).</summary>
    private const int MaxChatMembers = 500;

    /// <summary>Page cap for teams_list_chat_members so empty or pathological
    /// continuation pages cannot loop forever even if MaxChatMembers is never
    /// reached.</summary>
    private const int MaxChatMemberPages = 20;

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

    public async Task<IReadOnlyList<ChatInfo>> ListChatsAsync(
        int top = 25,
        int skip = 0,
        string? chatType = null,
        string? orderBy = null,
        CancellationToken ct = default)
    {
        int take = Math.Clamp(top, 1, 50);
        int offset = Math.Clamp(skip, 0, MaxScannedChats - 1);
        string? normalizedType = NormalizeChatType(chatType);
        bool orderByLastMessage = ParseChatOrderBy(orderBy);

        // Request server-side ordering by lastMessagePreview/createdDateTime desc
        // (and optional $filter=chatType eq '...') so recently active chats
        // arrive on page 1 even in mailboxes with >500 chats. When the tenant
        // rejects $orderby/$filter, or when 'lastUpdated' is requested, we scan
        // up to MaxScannedChats and filter/sort locally.
        var chats = new List<Chat>();
        (var page, bool serverOrdered, bool serverFiltered) =
            await GetChatsPageAsync(normalizedType, orderByLastMessage, ct).ConfigureAwait(false);
        int targetCount = offset + take;
        int pages = 0;
        while (page is not null)
        {
            pages++;
            chats.AddRange(page.Value ?? []);
            // Stop *before* the next request: no continuation left, server-ordered
            // + server-filtered results already cover offset + take, or the budget
            // for MaxScannedChats chats is spent – never fetch a page whose result
            // would be discarded. An empty page that still carries a continuation
            // is followed (the page budget bounds it, so an empty-page stream
            // cannot loop forever).
            if (string.IsNullOrWhiteSpace(page.OdataNextLink)
                || (serverOrdered && serverFiltered && chats.Count >= targetCount)
                || pages * ChatPageSize >= MaxScannedChats)
            {
                break;
            }

            page = await _client.Me.Chats
                .WithUrl(page.OdataNextLink)
                .GetAsync(cancellationToken: ct)
                .ConfigureAwait(false);
        }

        IEnumerable<Chat> filtered = normalizedType is not null && !serverFiltered
            ? chats.Where(c => string.Equals(c.ChatType?.ToString(), normalizedType, StringComparison.OrdinalIgnoreCase))
            : chats;

        IEnumerable<Chat> ordered = serverOrdered
            ? filtered
            : orderByLastMessage
                ? filtered.OrderByDescending(c => TeamsMapper.LatestActivity(c) ?? DateTimeOffset.MinValue)
                : filtered.OrderByDescending(c => c.LastUpdatedDateTime ?? DateTimeOffset.MinValue);

        return [.. ordered
            .Skip(offset)
            .Take(take)
            .Select(TeamsMapper.MapChat)];
    }

    private static string? NormalizeChatType(string? chatType)
    {
        if (string.IsNullOrWhiteSpace(chatType))
        {
            return null;
        }

        return chatType.Trim().ToLowerInvariant() switch
        {
            "oneonone" => "oneOnOne",
            "group" => "group",
            "meeting" => "meeting",
            _ => throw GraphServiceException.InvalidRequest(
                $"Unsupported chatType '{chatType}'.",
                "use 'oneOnOne', 'group', or 'meeting', or omit chatType")
        };
    }

    /// <summary>Returns true for 'lastMessage' (default), false for 'lastUpdated'.</summary>
    private static bool ParseChatOrderBy(string? orderBy)
    {
        if (string.IsNullOrWhiteSpace(orderBy))
        {
            return true;
        }

        return orderBy.Trim().ToLowerInvariant() switch
        {
            "lastmessage" => true,
            "lastupdated" => false,
            _ => throw GraphServiceException.InvalidRequest(
                $"Unsupported orderBy '{orderBy}'.",
                "use 'lastMessage' or 'lastUpdated', or omit orderBy")
        };
    }

    /// <summary>First page with server-side newest-message ordering, optional
    /// chatType filter and $expand=members,lastMessagePreview. Falls back on
    /// HTTP 400 so tenants that reject an expand, $filter or $orderby still get
    /// a working listing.</summary>
    private async Task<(ChatCollectionResponse? Page, bool ServerOrdered, bool ServerFiltered)> GetChatsPageAsync(
        string? chatType, bool orderByLastMessage, CancellationToken ct)
    {
        string? filter = chatType is null ? null : $"chatType eq '{chatType}'";
        string[]? serverOrder = orderByLastMessage ? ["lastMessagePreview/createdDateTime desc"] : null;

        var attempts = new List<(string[]? Expand, string[]? Orderby, string? Filter, bool ServerOrdered, bool ServerFiltered)>();
        if (serverOrder is not null)
        {
            if (filter is not null)
            {
                attempts.Add((["members", "lastMessagePreview"], serverOrder, filter, true, true));
                attempts.Add((["lastMessagePreview"], serverOrder, filter, true, true));
            }

            attempts.Add((["members", "lastMessagePreview"], serverOrder, null, true, false));
            attempts.Add((["lastMessagePreview"], serverOrder, null, true, false));
        }

        if (filter is not null)
        {
            attempts.Add((["members", "lastMessagePreview"], null, filter, false, true));
            attempts.Add((["members"], null, filter, false, true));
        }

        attempts.Add((["members", "lastMessagePreview"], null, null, false, false));
        attempts.Add((["members"], null, null, false, false));
        attempts.Add((null, null, null, false, false));

        for (int i = 0; i < attempts.Count; i++)
        {
            var (expand, orderby, currentFilter, serverOrdered, serverFiltered) = attempts[i];
            try
            {
                var page = await _client.Me.Chats.GetAsync(c =>
                {
                    c.QueryParameters.Top = ChatPageSize;
                    if (expand is not null)
                    {
                        c.QueryParameters.Expand = expand;
                    }

                    if (orderby is not null)
                    {
                        c.QueryParameters.Orderby = orderby;
                    }

                    if (currentFilter is not null)
                    {
                        c.QueryParameters.Filter = currentFilter;
                    }
                }, ct).ConfigureAwait(false);
                return (page, serverOrdered, serverFiltered || chatType is null);
            }
            catch (ApiException ex) when (ex.ResponseStatusCode == 400 && i < attempts.Count - 1)
            {
                // Try the next less demanding query shape.
            }
        }

        return (null, false, chatType is null);
    }

    /// <summary>Members of one chat. The endpoint documents no OData query
    /// parameters, so only Graph's continuation link (if any) is followed,
    /// bounded by MaxChatMembers and MaxChatMemberPages.</summary>
    public async Task<IReadOnlyList<ChatMemberInfo>> ListChatMembersAsync(
        string chatId, CancellationToken ct = default)
    {
        RequireId(chatId, "chatId", "call teams_list_chats to get valid chat ids");
        string id = chatId.Trim();
        var members = new List<ConversationMember>();
        var page = await _client.Me.Chats[id].Members
            .GetAsync(cancellationToken: ct).ConfigureAwait(false);
        int pages = 0;
        while (page is not null)
        {
            pages++;
            if (page.Value is { Count: > 0 } items)
            {
                int remaining = MaxChatMembers - members.Count;
                if (items.Count <= remaining)
                {
                    members.AddRange(items);
                }
                else
                {
                    members.AddRange(items.Take(remaining));
                }
            }

            if (members.Count >= MaxChatMembers
                || pages >= MaxChatMemberPages
                || string.IsNullOrWhiteSpace(page.OdataNextLink))
            {
                break;
            }

            page = await _client.Me.Chats[id].Members
                .WithUrl(page.OdataNextLink)
                .GetAsync(cancellationToken: ct)
                .ConfigureAwait(false);
        }

        return [.. members.Select(TeamsMapper.MapChatMember)];
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

    public async Task<MessageSearchResult> SearchMessagesAsync(
        string query, int top = 25, CancellationToken ct = default)
    {
        RequireId(query, "query", "provide a search term or KQL query (e.g. 'budget' or 'from:alice@contoso.com')");
        int size = Math.Clamp(top, 1, 50);
        string payload = JsonSerializer.Serialize(new
        {
            requests = new[]
            {
                new
                {
                    entityTypes = new[] { "chatMessage" },
                    query = new { queryString = query.Trim() },
                    from = 0,
                    size
                }
            }
        });
        using JsonDocument document = await PostJsonAsync("/search/query", payload, ct).ConfigureAwait(false);
        return TeamsMapper.MapSearchResult(document.RootElement);
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

    private async Task<JsonDocument> PostJsonAsync(string path, string jsonBody, CancellationToken ct)
    {
        var adapter = _client.RequestAdapter
            ?? throw GraphServiceException.GraphError(0, null, "Graph request adapter is unavailable.");
        string baseUrl = adapter.BaseUrl?.TrimEnd('/')
            ?? throw GraphServiceException.GraphError(0, null, "Graph request adapter has no base URL.");
        var request = new RequestInformation
        {
            HttpMethod = Method.POST,
            UrlTemplate = "{+baseurl}" + path,
            PathParameters = new Dictionary<string, object>
            {
                ["baseurl"] = baseUrl
            }
        };
        request.Headers.Add("Accept", "application/json");
        request.SetStreamContent(new MemoryStream(Encoding.UTF8.GetBytes(jsonBody)), "application/json");
        using Stream? stream = await adapter.SendPrimitiveAsync<Stream>(
            request,
            new Dictionary<string, ParsableFactory<IParsable>>(),
            ct).ConfigureAwait(false);
        if (stream is null)
        {
            throw GraphServiceException.GraphError(200, null, "Graph returned an empty search response.");
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
