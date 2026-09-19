namespace MicrosoftMcp.Teams;

/// <summary>Teams access for the signed-in user with guarded message sending.</summary>
public interface IGraphTeamsService
{
    Task<IReadOnlyList<TeamInfo>> ListTeamsAsync(CancellationToken ct = default);
    Task<IReadOnlyList<ChannelInfo>> ListChannelsAsync(string teamId, CancellationToken ct = default);
    Task<IReadOnlyList<MessageSummary>> ListChannelMessagesAsync(
        string teamId, string channelId, int top = 25, CancellationToken ct = default);
    Task<IReadOnlyList<MessageSummary>> ListMessageRepliesAsync(
        string teamId, string channelId, string messageId, int top = 25, CancellationToken ct = default);
    Task<MessageDetail> ReadChannelMessageAsync(
        string teamId, string channelId, string messageId, CancellationToken ct = default);
    Task<IReadOnlyList<ChatInfo>> ListChatsAsync(int top = 25, CancellationToken ct = default);
    Task<IReadOnlyList<MessageSummary>> ListChatMessagesAsync(
        string chatId, int top = 25, CancellationToken ct = default);
    Task<MessageDetail> ReadChatMessageAsync(
        string chatId, string messageId, CancellationToken ct = default);
    Task<MessageDetail> SendChannelMessageAsync(
        string teamId, string channelId, string body, CancellationToken ct = default);
    Task<MessageDetail> SendChatMessageAsync(
        string chatId, string body, CancellationToken ct = default);
    Task<IReadOnlyList<MeetingTranscriptInfo>> ListMeetingTranscriptsAsync(
        string meetingId, int top = 25, CancellationToken ct = default);
    Task<MeetingTranscriptDetail> ReadMeetingTranscriptAsync(
        string meetingId, string transcriptId, CancellationToken ct = default);
    Task<IReadOnlyList<MeetingInsightInfo>> ListMeetingInsightsAsync(
        string meetingId, int top = 25, CancellationToken ct = default);
    Task<MeetingInsightDetail> ReadMeetingInsightAsync(
        string meetingId, string insightId, CancellationToken ct = default);
}
