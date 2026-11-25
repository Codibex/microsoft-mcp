namespace MicrosoftMcp.Teams;

public sealed record TeamInfo(
    string Id,
    string DisplayName,
    string? Description,
    string? Visibility,
    bool IsArchived);

public sealed record ChannelInfo(
    string Id,
    string DisplayName,
    string? Description,
    string? MembershipType);

/// <summary>SharePoint file storage behind a channel. The driveId/folderId
/// pair plugs directly into the sharepoint_* tools (explicit-drive path),
/// keeping the fixed-drive onedrive_* tools untouched.</summary>
public sealed record ChannelFilesFolderInfo(
    string TeamId,
    string ChannelId,
    string? DriveId,
    string FolderId,
    string? FolderName,
    string? SiteId,
    string? WebUrl);

/// <summary>Chat listing entry. Members are the chat's participants
/// (from $expand=members), so a 1:1 chat names its counterparty without
/// reading messages; empty when the expand was unavailable.</summary>
public sealed record ChatInfo(
    string Id,
    string? Topic,
    string? ChatType,
    DateTimeOffset? LastUpdated,
    IReadOnlyList<ChatMemberInfo>? Members = null);

public sealed record ChatMemberInfo(
    string? DisplayName,
    string? Email,
    string? UserId,
    IReadOnlyList<string>? Roles = null);

public sealed record MeetingTranscriptInfo(
    string Id,
    string? MeetingId,
    string? CallId,
    DateTimeOffset? Created,
    DateTimeOffset? Ended,
    string? ContentCorrelationId);

public sealed record MeetingTranscriptDetail(
    string Id,
    string? MeetingId,
    string? CallId,
    DateTimeOffset? Created,
    DateTimeOffset? Ended,
    string? ContentCorrelationId,
    string Content,
    string ContentType);

public sealed record MeetingInsightInfo(
    string Id,
    string? CallId,
    string? ContentCorrelationId,
    DateTimeOffset? Created,
    DateTimeOffset? Ended);

public sealed record MeetingInsightDetail(
    string Id,
    string? CallId,
    string? ContentCorrelationId,
    DateTimeOffset? Created,
    DateTimeOffset? Ended,
    IReadOnlyList<MeetingNoteInfo> MeetingNotes,
    IReadOnlyList<MeetingActionItemInfo> ActionItems,
    IReadOnlyList<MeetingMentionInfo> Mentions);

public sealed record MeetingNoteInfo(
    string? Title,
    string? Text,
    IReadOnlyList<MeetingNoteSubpointInfo> Subpoints);

public sealed record MeetingNoteSubpointInfo(
    string? Title,
    string? Text);

public sealed record MeetingActionItemInfo(
    string? Title,
    string? Text,
    string? OwnerDisplayName);

public sealed record MeetingMentionInfo(
    DateTimeOffset? EventTime,
    string? TranscriptUtterance,
    string? Speaker);

public sealed record ReactionDto(string? Type, string? User, DateTimeOffset? Created);

public sealed record MentionDto(string? Text, string? Mentioned);

public sealed record MessageSummary(
    string Id,
    string? MessageType,
    string? From,
    DateTimeOffset? Created,
    string? Preview,
    string? ContentType,
    int ReactionCount,
    int AttachmentCount,
    string? ReplyToId,
    string? WebUrl);

public sealed record MessageDetail(
    string Id,
    string? MessageType,
    string? From,
    DateTimeOffset? Created,
    string? Content,
    string? ContentType,
    IReadOnlyList<ReactionDto> Reactions,
    IReadOnlyList<MentionDto> Mentions,
    int AttachmentCount,
    string? ReplyToId,
    string? WebUrl);

/// <summary>One Teams message hit from Microsoft Search (chat or channel).
/// ChatId + MessageId plug into teams_read_chat_message; TeamId/ChannelId +
/// MessageId plug into teams_read_channel_message for channel root messages
/// (channel reply hits cannot be opened via the root-message endpoint — use
/// WebUrl when available).</summary>
public sealed record MessageSearchHit(
    string MessageId,
    string? ChatId,
    string? TeamId,
    string? ChannelId,
    string? From,
    string? FromAddress,
    DateTimeOffset? Created,
    string? Summary,
    string? WebUrl);

public sealed record MessageSearchResult(
    IReadOnlyList<MessageSearchHit> Hits,
    int Total,
    bool MoreResultsAvailable);
