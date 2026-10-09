using Microsoft.Graph.Models;

namespace MicrosoftMcp.Planner;

internal static class PlannerMapper
{
    internal static PlannerPlanInfo MapPlan(PlannerPlan plan) => new(
        plan.Id ?? string.Empty,
        plan.Title,
        plan.Owner,
        plan.CreatedDateTime);

    internal static PlannerTaskSummary MapSummary(PlannerTask task) => new(
        task.Id ?? string.Empty,
        task.PlanId,
        task.BucketId,
        task.Title,
        task.PercentComplete,
        task.Priority,
        FormatDto(task.DueDateTime),
        task.CreatedDateTime,
        Assignees(task));

    internal static PlannerTaskDetail MapDetail(
        PlannerTask task, PlannerTaskDetails? details, string? bucketName) => new(
        task.Id ?? string.Empty,
        task.PlanId,
        task.BucketId,
        bucketName,
        task.Title,
        task.PercentComplete,
        task.Priority,
        FormatDto(task.DueDateTime),
        FormatDto(task.StartDateTime),
        task.CreatedDateTime,
        Assignees(task),
        details?.Description?.Trim(),
        Checklist(details));

    internal static IReadOnlyList<string> Assignees(PlannerTask task) =>
        task.Assignments?.AdditionalData is null
            ? []
            : [.. task.Assignments.AdditionalData.Keys.OrderBy(k => k, StringComparer.Ordinal)];

    internal static IReadOnlyList<PlannerChecklistEntry> Checklist(PlannerTaskDetails? details) =>
        details?.Checklist?.AdditionalData is null
            ? []
            : [.. details.Checklist.AdditionalData.Values
                .OfType<PlannerChecklistItem>()
                .OrderBy(c => c.OrderHint, StringComparer.Ordinal)
                .Select(c => new PlannerChecklistEntry(c.Title, c.IsChecked))];

    internal static string? FormatDto(DateTimeOffset? dt) =>
        dt is null ? null : dt.Value.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ss'Z'");
}
