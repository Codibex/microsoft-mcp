namespace MicrosoftMcp.Teams.Tests;

using MicrosoftMcp.Common;

/// <summary>In-memory fake of <see cref="IGraphTeamsService"/>. No Graph, no network.</summary>
internal sealed class FakeGraphTeamsService : IGraphTeamsService
{
    private sealed class StoredMessage
    {
        public required string Id { get; init; }
        public required string OwnerId { get; init; }
        public string? ParentId { get; init; }
        public string Subject { get; init; } = string.Empty;
        public string Body { get; init; } = string.Empty;
        public string From { get; init; } = "Alice";
    }

    private readonly Dictionary<string, TeamInfo> _teams = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, (string TeamId, ChannelInfo Channel)> _channels = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, ChatInfo> _chats = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, StoredMessage> _messages = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, MeetingTranscriptInfo> _transcripts = new(StringComparer.OrdinalIgnoreCase)
    {
        ["tr-1"] = new("tr-1", "meeting-1", "call-1", DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow, "corr-1")
    };
    private readonly Dictionary<string, MeetingInsightDetail> _insights = new(StringComparer.OrdinalIgnoreCase)
    {
        ["insight-1"] = new(
            "insight-1", "call-1", "corr-1", DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow,
            [new MeetingNoteInfo("Decisions", "The team agreed on the plan.",
                [new MeetingNoteSubpointInfo("Next step", "Publish the draft.")])],
            [new MeetingActionItemInfo("Publish draft", "Publish the meeting draft.", "Alice")],
            [new MeetingMentionInfo(DateTimeOffset.UtcNow.AddMinutes(-2), "Alice mentioned Bob.", "Alice")])
    };

    public FakeGraphTeamsService()
    {
        _teams["t-eng"] = new TeamInfo("t-eng", "Engineering", null, "private", false);
        _channels["c-general"] = ("t-eng", new ChannelInfo("c-general", "General", null, "standard"));
        _channels["c-random"] = ("t-eng", new ChannelInfo("c-random", "Random", null, "standard"));
        _chats["chat-1"] = new ChatInfo("chat-1", null, "oneOnOne", DateTimeOffset.UtcNow);

        Add(new StoredMessage { Id = "m-hello", OwnerId = "c-general", Subject = "Hallo", Body = "<p>Willkommen!</p>" });
        Add(new StoredMessage { Id = "m-r1", OwnerId = "c-general", ParentId = "m-hello", Subject = "Re: Hallo", Body = "Danke!" });
        Add(new StoredMessage { Id = "m-other", OwnerId = "c-random", Subject = "Anderes", Body = "Bla" });
        Add(new StoredMessage { Id = "cm-1", OwnerId = "chat-1", Subject = "Hi", Body = "Kurze Frage" });
    }

    private void Add(StoredMessage m) => _messages[m.Id] = m;

    private static MessageSummary ToSummary(StoredMessage m) => new(
        m.Id, "message", m.From, DateTimeOffset.UtcNow,
        m.Subject, "html", 0, 0, m.ParentId, null);

    private static MessageDetail ToDetail(StoredMessage m) => new(
        m.Id, "message", m.From, DateTimeOffset.UtcNow, m.Body, "html",
        [], [], 0, m.ParentId, null);

    private static void Require(string value, string what, string hint)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw GraphServiceException.InvalidRequest($"{what} must not be empty.", hint);
        }
    }

    public Task<IReadOnlyList<TeamInfo>> ListTeamsAsync(CancellationToken ct = default) =>
        Task.FromResult<IReadOnlyList<TeamInfo>>([.. _teams.Values]);

    public Task<IReadOnlyList<ChannelInfo>> ListChannelsAsync(string teamId, CancellationToken ct = default)
    {
        Require(teamId, "teamId", "call teams_list_teams to get valid team ids");
        if (!_teams.ContainsKey(teamId.Trim()))
        {
            throw GraphServiceException.TeamNotFound(teamId, "fake");
        }

        return Task.FromResult<IReadOnlyList<ChannelInfo>>(
            [.. _channels.Values.Where(c => c.TeamId == teamId.Trim()).Select(c => c.Channel)]);
    }

    private IReadOnlyList<StoredMessage> Thread(string ownerId, string? parentId, int top) =>
        [.. _messages.Values
            .Where(m => m.OwnerId == ownerId && m.ParentId == parentId)
            .Take(Math.Clamp(top, 1, 100))];

    public Task<IReadOnlyList<MessageSummary>> ListChannelMessagesAsync(
        string teamId, string channelId, int top = 25, CancellationToken ct = default)
    {
        Require(teamId, "teamId", "call teams_list_teams to get valid team ids");
        Require(channelId, "channelId", "call teams_list_channels for the team to get valid channel ids");
        if (!_channels.TryGetValue(channelId.Trim(), out var ch) || ch.TeamId != teamId.Trim())
        {
            throw GraphServiceException.ChannelNotFound(channelId, "fake");
        }

        return Task.FromResult<IReadOnlyList<MessageSummary>>(
            [.. Thread(ch.Channel.Id, null, top).Select(ToSummary)]);
    }

    public Task<IReadOnlyList<MessageSummary>> ListMessageRepliesAsync(
        string teamId, string channelId, string messageId, int top = 25, CancellationToken ct = default)
    {
        Require(messageId, "messageId", "call teams_list_channel_messages to get valid message ids");
        return Task.FromResult<IReadOnlyList<MessageSummary>>(
            [.. Thread(channelId, messageId.Trim(), top).Select(ToSummary)]);
    }

    public Task<MessageDetail> ReadChannelMessageAsync(
        string teamId, string channelId, string messageId, CancellationToken ct = default)
    {
        Require(messageId, "messageId", "call teams_list_channel_messages to get valid message ids");
        return Task.FromResult(
            _messages.TryGetValue(messageId.Trim(), out var m)
                ? ToDetail(m)
                : throw GraphServiceException.TeamsMessageNotFound(messageId, "fake"));
    }

    public Task<IReadOnlyList<ChatInfo>> ListChatsAsync(int top = 25, CancellationToken ct = default) =>
        Task.FromResult<IReadOnlyList<ChatInfo>>(
            [.. _chats.Values.Take(Math.Clamp(top, 1, 100))]);

    public Task<IReadOnlyList<MessageSummary>> ListChatMessagesAsync(
        string chatId, int top = 25, CancellationToken ct = default)
    {
        Require(chatId, "chatId", "call teams_list_chats to get valid chat ids");
        if (!_chats.ContainsKey(chatId.Trim()))
        {
            throw GraphServiceException.ChatNotFound(chatId, "fake");
        }

        return Task.FromResult<IReadOnlyList<MessageSummary>>(
            [.. Thread(chatId.Trim(), null, top).Select(ToSummary)]);
    }

    public Task<IReadOnlyList<MessageSummary>> ListChatRepliesAsync(
        string chatId, string messageId, int top = 25, CancellationToken ct = default)
    {
        Require(messageId, "messageId", "call teams_list_chat_messages to get valid message ids");
        return Task.FromResult<IReadOnlyList<MessageSummary>>(
            [.. Thread(chatId, messageId.Trim(), top).Select(ToSummary)]);
    }

    public Task<MessageDetail> ReadChatMessageAsync(
        string chatId, string messageId, CancellationToken ct = default)
    {
        Require(messageId, "messageId", "call teams_list_chat_messages to get valid message ids");
        return Task.FromResult(
            _messages.TryGetValue(messageId.Trim(), out var m)
                ? ToDetail(m)
                : throw GraphServiceException.TeamsMessageNotFound(messageId, "fake"));
    }

    public Task<IReadOnlyList<MeetingTranscriptInfo>> ListMeetingTranscriptsAsync(
        string meetingId, int top = 25, CancellationToken ct = default)
    {
        Require(meetingId, "meetingId", "use the online meeting id from Microsoft Graph");
        return Task.FromResult<IReadOnlyList<MeetingTranscriptInfo>>(
            [.. _transcripts.Values.Where(t => t.MeetingId == meetingId.Trim()).Take(Math.Clamp(top, 1, 100))]);
    }

    public Task<MeetingTranscriptDetail> ReadMeetingTranscriptAsync(
        string meetingId, string transcriptId, CancellationToken ct = default)
    {
        Require(meetingId, "meetingId", "use the online meeting id from Microsoft Graph");
        Require(transcriptId, "transcriptId", "call teams_list_meeting_transcripts to get valid transcript ids");
        if (!_transcripts.TryGetValue(transcriptId.Trim(), out var transcript)
            || transcript.MeetingId != meetingId.Trim())
        {
            throw GraphServiceException.MeetingTranscriptNotFound(transcriptId, "fake");
        }

        return Task.FromResult(new MeetingTranscriptDetail(
            transcript.Id, transcript.MeetingId, transcript.CallId, transcript.Created,
            transcript.Ended, transcript.ContentCorrelationId, "WEBVTT\n\nHello", "text/vtt"));
    }

    public Task<IReadOnlyList<MeetingInsightInfo>> ListMeetingInsightsAsync(
        string meetingId, int top = 25, CancellationToken ct = default)
    {
        Require(meetingId, "meetingId", "use the online meeting id from Microsoft Graph");
        return Task.FromResult<IReadOnlyList<MeetingInsightInfo>>(
            [.. _insights.Values
                .Where(i => _transcripts.Values.Any(t => t.MeetingId == meetingId.Trim() && t.CallId == i.CallId))
                .Take(Math.Clamp(top, 1, 100))
                .Select(i => new MeetingInsightInfo(i.Id, i.CallId, i.ContentCorrelationId, i.Created, i.Ended))]);
    }

    public Task<MeetingInsightDetail> ReadMeetingInsightAsync(
        string meetingId, string insightId, CancellationToken ct = default)
    {
        Require(meetingId, "meetingId", "use the online meeting id from Microsoft Graph");
        Require(insightId, "insightId", "call teams_list_meeting_insights to get valid insight ids");
        if (!_insights.TryGetValue(insightId.Trim(), out var insight)
            || !_transcripts.Values.Any(t => t.MeetingId == meetingId.Trim() && t.CallId == insight.CallId))
        {
            throw GraphServiceException.MeetingInsightNotFound(insightId, "fake");
        }

        return Task.FromResult(insight);
    }
}
