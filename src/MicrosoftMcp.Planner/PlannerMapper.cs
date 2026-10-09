using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Serialization;

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
        [.. AllChecklistEntries(details).OrderBy(e => e.Order, StringComparer.Ordinal).Select(e => e.Entry)];

    private static IEnumerable<(string? Order, PlannerChecklistEntry Entry)> AllChecklistEntries(
        PlannerTaskDetails? details)
    {
        if (details?.Checklist?.AdditionalData is null)
        {
            yield break;
        }

        foreach (object value in details.Checklist.AdditionalData.Values)
        {
            switch (value)
            {
                case PlannerChecklistItem item:
                    yield return (item.OrderHint, new PlannerChecklistEntry(item.Title, item.IsChecked));
                    break;
                // Real Graph responses carry no discriminator for these
                // open-type entries, so Kiota materializes them as
                // UntypedObject (verified against live-shaped JSON).
                case UntypedObject obj when TryReadChecklist(obj, out var entry, out var order):
                    yield return (order, entry);
                    break;
            }
        }
    }

    private static bool TryReadChecklist(
        UntypedObject obj, out PlannerChecklistEntry entry, out string? order)
    {
        IDictionary<string, UntypedNode> props = obj.GetValue();
        // NOTE: UntypedNode.GetValue() itself throws NotImplementedException;
        // values must be read through the concrete node types.
        string? title = props.TryGetValue("title", out UntypedNode? titleNode)
            && titleNode is UntypedString titleStr
            ? titleStr.GetValue()
            : null;
        bool? isChecked = props.TryGetValue("isChecked", out UntypedNode? checkedNode)
            && checkedNode is UntypedBoolean checkedBool
            ? checkedBool.GetValue()
            : null;
        order = props.TryGetValue("orderHint", out UntypedNode? orderNode)
            && orderNode is UntypedString orderStr
            ? orderStr.GetValue()
            : null;
        if (title is null && isChecked is null)
        {
            entry = new PlannerChecklistEntry(null, null);
            return false;
        }

        entry = new PlannerChecklistEntry(title, isChecked);
        return true;
    }

    internal static string? FormatDto(DateTimeOffset? dt) =>
        dt is null ? null : dt.Value.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ss'Z'");
}
