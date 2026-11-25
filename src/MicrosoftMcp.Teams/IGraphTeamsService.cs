namespace MicrosoftMcp.Teams;

/// <summary>Read-only Teams access for the signed-in user.</summary>
public interface IGraphTeamsService
{
    Task<IReadOnlyList<TeamInfo>> ListTeamsAsync(CancellationToken ct = default);
    Task<IReadOnlyList<ChannelInfo>> ListChannelsAsync(string teamId, CancellationToken ct = default);
    /// <summary>Resolves the SharePoint files folder behind a channel.
    /// Bridge into the sharepoint_* tools: the returned driveId/folderId
    /// address the channel library on its explicit-drive path.</summary>
    Task<ChannelFilesFolderInfo> GetChannelFilesFolderAsync(
        string teamId, string channelId, CancellationToken ct = default);
    Task<IReadOnlyList<MessageSummary>> ListChannelMessagesAsync(
        string teamId, string channelId, int top = 25, CancellationToken ct = default);
    Task<IReadOnlyList<MessageSummary>> ListMessageRepliesAsync(
        string teamId, string channelId, string messageId, int top = 25, CancellationToken ct = default);
    Task<MessageDetail> ReadChannelMessageAsync(
        string teamId, string channelId, string messageId, CancellationToken ct = default);
    Task<IReadOnlyList<ChatInfo>> ListChatsAsync(
        int top = 25,
        int skip = 0,
        string? chatType = null,
        string? orderBy = null,
        CancellationToken ct = default);
    /// <summary>Members of one chat (displayName/email/userId/roles).
    /// Chat.ReadBasic is enough; needed for group chats beyond the
    /// 25-member $expand=members cap of teams_list_chats.</summary>
    Task<IReadOnlyList<ChatMemberInfo>> ListChatMembersAsync(
        string chatId, CancellationToken ct = default);
    Task<IReadOnlyList<MessageSummary>> ListChatMessagesAsync(
        string chatId, int top = 25, CancellationToken ct = default);
    Task<MessageDetail> ReadChatMessageAsync(
        string chatId, string messageId, CancellationToken ct = default);
    /// <summary>Search Teams messages across chats and channels via Microsoft
    /// Search (POST /v1.0/search/query with entityTypes: ["chatMessage"]).
    /// Work/school delegated accounts only.</summary>
    Task<MessageSearchResult> SearchMessagesAsync(
        string query, int top = 25, CancellationToken ct = default);
    Task<IReadOnlyList<MeetingTranscriptInfo>> ListMeetingTranscriptsAsync(
        string meetingId, int top = 25, CancellationToken ct = default);
    Task<MeetingTranscriptDetail> ReadMeetingTranscriptAsync(
        string meetingId, string transcriptId, CancellationToken ct = default);
    Task<IReadOnlyList<MeetingInsightInfo>> ListMeetingInsightsAsync(
        string meetingId, int top = 25, CancellationToken ct = default);
    Task<MeetingInsightDetail> ReadMeetingInsightAsync(
        string meetingId, string insightId, CancellationToken ct = default);
}
