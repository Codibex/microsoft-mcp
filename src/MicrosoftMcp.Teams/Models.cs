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

public sealed record ChatInfo(
    string Id,
    string? Topic,
    string? ChatType,
    DateTimeOffset? LastUpdated);

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
