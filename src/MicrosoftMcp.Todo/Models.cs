namespace MicrosoftMcp.Todo;

public sealed record TodoListInfo(
    string Id,
    string DisplayName,
    bool? IsOwner,
    bool? IsShared,
    string? WellknownListName);

public sealed record TodoTaskSummary(
    string Id,
    string ListId,
    string? Title,
    string? Status,
    string? Importance,
    string? DueDateTime,
    DateTimeOffset? Created,
    DateTimeOffset? LastModified,
    bool? HasAttachments);

public sealed record TodoTaskDetail(
    string Id,
    string ListId,
    string? Title,
    string? Status,
    string? Importance,
    string? DueDateTime,
    string? ReminderDateTime,
    string? CompletedDateTime,
    DateTimeOffset? Created,
    DateTimeOffset? LastModified,
    string? Body,
    string? BodyContentType,
    IReadOnlyList<string> Categories);
