namespace MicrosoftMcp.Planner;

/// <summary>Read-only Planner access for the signed-in user: tasks assigned
/// to me, plans of a group, tasks of a plan, single task with details.</summary>
public interface IGraphPlannerService
{
    Task<IReadOnlyList<PlannerTaskSummary>> ListMyTasksAsync(int top = 50, CancellationToken ct = default);
    Task<IReadOnlyList<PlannerPlanInfo>> ListPlansAsync(string groupId, int top = 50, CancellationToken ct = default);
    Task<IReadOnlyList<PlannerTaskSummary>> ListPlanTasksAsync(string planId, int top = 50, CancellationToken ct = default);
    Task<PlannerTaskDetail> ReadTaskAsync(string taskId, CancellationToken ct = default);
}
