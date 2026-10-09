namespace MicrosoftMcp.Planner.Tests;

using MicrosoftMcp.Common;

/// <summary>In-memory fake of <see cref="IGraphPlannerService"/>. No Graph, no network.</summary>
internal sealed class FakeGraphPlannerService : IGraphPlannerService
{
    private readonly Dictionary<string, (string GroupId, PlannerPlanInfo Plan)> _plans = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, PlannerTaskDetail> _tasks = new(StringComparer.OrdinalIgnoreCase);

    public FakeGraphPlannerService()
    {
        _plans["plan-1"] = ("g-eng", new PlannerPlanInfo("plan-1", "Launch", "g-eng", DateTimeOffset.UtcNow));
        _plans["plan-2"] = ("g-eng", new PlannerPlanInfo("plan-2", "Docs", "g-eng", DateTimeOffset.UtcNow));
        Add(new PlannerTaskDetail(
            "task-1", "plan-1", "bucket-1", "To do", "Offer schreiben", 0, 5,
            "2026-10-15T17:00:00 (UTC)", null, DateTimeOffset.UtcNow,
            ["user-alice"], "Angebot an Contoso", [new PlannerChecklistEntry("Entwurf", false)]));
        Add(new PlannerTaskDetail(
            "task-2", "plan-1", "bucket-2", "Done", "Review", 100, 5,
            null, null, DateTimeOffset.UtcNow,
            ["user-bob"], null, []));
    }

    private void Add(PlannerTaskDetail t) => _tasks[t.Id] = t;

    private static void Require(string value, string what, string hint)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw GraphServiceException.InvalidRequest($"{what} must not be empty.", hint);
        }
    }

    private static PlannerTaskSummary ToSummary(PlannerTaskDetail t) => new(
        t.Id, t.PlanId, t.BucketId, t.Title, t.PercentComplete, t.Priority,
        t.DueDateTime, t.Created, t.AssigneeIds);

    public Task<IReadOnlyList<PlannerTaskSummary>> ListMyTasksAsync(int top = 50, CancellationToken ct = default) =>
        Task.FromResult<IReadOnlyList<PlannerTaskSummary>>(
            [.. _tasks.Values
                .Where(t => t.AssigneeIds.Contains("user-alice"))
                .Take(Math.Clamp(top, 1, 100))
                .Select(ToSummary)]);

    public Task<IReadOnlyList<PlannerPlanInfo>> ListPlansAsync(string groupId, int top = 50, CancellationToken ct = default)
    {
        Require(groupId, "groupId", "use the Microsoft 365 group id backing the plan");
        if (!string.Equals(groupId.Trim(), "g-eng", StringComparison.OrdinalIgnoreCase))
        {
            throw GraphServiceException.GroupNotFound(groupId, "fake");
        }

        return Task.FromResult<IReadOnlyList<PlannerPlanInfo>>(
            [.. _plans.Values.Select(p => p.Plan).Take(Math.Clamp(top, 1, 100))]);
    }

    public Task<IReadOnlyList<PlannerTaskSummary>> ListPlanTasksAsync(string planId, int top = 50, CancellationToken ct = default)
    {
        Require(planId, "planId", "call planner_list_plans with the group id to get valid plan ids");
        if (!_plans.ContainsKey(planId.Trim()))
        {
            throw GraphServiceException.PlannerPlanNotFound(planId, "fake");
        }

        return Task.FromResult<IReadOnlyList<PlannerTaskSummary>>(
            [.. _tasks.Values
                .Where(t => t.PlanId == planId.Trim())
                .Take(Math.Clamp(top, 1, 100))
                .Select(ToSummary)]);
    }

    public Task<PlannerTaskDetail> ReadTaskAsync(string taskId, CancellationToken ct = default)
    {
        Require(taskId, "taskId", "call planner_list_my_tasks or planner_list_plan_tasks to get valid task ids");
        return Task.FromResult(
            _tasks.TryGetValue(taskId.Trim(), out var task)
                ? task
                : throw GraphServiceException.PlannerTaskNotFound(taskId, "fake"));
    }
}
