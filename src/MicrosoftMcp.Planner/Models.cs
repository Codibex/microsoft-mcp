namespace MicrosoftMcp.Planner;

public sealed record PlannerPlanInfo(
    string Id,
    string? Title,
    string? Owner,
    DateTimeOffset? Created);

public sealed record PlannerTaskSummary(
    string Id,
    string? PlanId,
    string? BucketId,
    string? Title,
    int? PercentComplete,
    int? Priority,
    string? DueDateTime,
    DateTimeOffset? Created,
    IReadOnlyList<string> AssigneeIds);

public sealed record PlannerChecklistEntry(
    string? Title,
    bool? IsChecked);

public sealed record PlannerTaskDetail(
    string Id,
    string? PlanId,
    string? BucketId,
    string? BucketName,
    string? Title,
    int? PercentComplete,
    int? Priority,
    string? DueDateTime,
    string? StartDateTime,
    DateTimeOffset? Created,
    IReadOnlyList<string> AssigneeIds,
    string? Description,
    IReadOnlyList<PlannerChecklistEntry> Checklist);
