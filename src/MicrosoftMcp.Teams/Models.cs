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

public sealed record ChatInfo(
    string Id,
    string? Topic,
    string? ChatType,
    DateTimeOffset? LastUpdated);

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
