using System.ComponentModel;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using MicrosoftMcp.Common;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace MicrosoftMcp.Planner;

public static class PlannerServiceRegistration
{
    public static IServiceCollection AddPlanner(this IServiceCollection services)
    {
        services.AddSingleton<IGraphPlannerService, GraphPlannerService>();
        return services;
    }
}

/// <summary>Planner: read-only view on commitments owned by teams
/// (who owns what). No writes offered.</summary>
[McpServerToolType]
public sealed class PlannerTools(IGraphPlannerService planner, ILogger<PlannerTools> log)
{
    private async Task<CallToolResult> InvokeAsync<T>(
        string operation, Func<Task<T>> call, string resource = "planner-task")
    {
        try
        {
            return ToolResult.Ok(await call().ConfigureAwait(false));
        }
        catch (GraphServiceException ex)
        {
            log.LogWarning(ex, "Tool {Operation} failed with {Code}", operation, ex.Code);
            return ToolResult.Fail(ex);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            var mapped = GraphErrorMapper.ToGraphServiceException(ex, operation, resource);
            log.LogError(ex, "Tool {Operation} failed unexpectedly ({Code})", operation, mapped.Code);
            return ToolResult.Fail(mapped);
        }
    }

    [McpServerTool, Description("List Planner tasks assigned to me. Read-only.")]
    public Task<CallToolResult> planner_list_my_tasks(
        [Description("Max tasks 1-100")] int top = 50,
        CancellationToken ct = default) =>
        InvokeAsync("planner_list_my_tasks", () => planner.ListMyTasksAsync(top, ct));

    [McpServerTool, Description("List Planner plans of a Microsoft 365 group. Read-only.")]
    public Task<CallToolResult> planner_list_plans(
        [Description("Microsoft 365 group id backing the plans")] string groupId,
        [Description("Max plans 1-100")] int top = 50,
        CancellationToken ct = default) =>
        InvokeAsync("planner_list_plans", () => planner.ListPlansAsync(groupId, top, ct), "group");

    [McpServerTool, Description("List tasks of a Planner plan. Read-only.")]
    public Task<CallToolResult> planner_list_plan_tasks(
        [Description("Plan id from planner_list_plans")] string planId,
        [Description("Max tasks 1-100")] int top = 50,
        CancellationToken ct = default) =>
        InvokeAsync("planner_list_plan_tasks", () => planner.ListPlanTasksAsync(planId, top, ct), "planner-plan");

    [McpServerTool, Description("Read a Planner task with description, checklist, assignees and bucket. Read-only.")]
    public Task<CallToolResult> planner_read_task(
        [Description("Task id from planner_list_my_tasks or planner_list_plan_tasks")] string taskId,
        CancellationToken ct = default) =>
        InvokeAsync("planner_read_task", () => planner.ReadTaskAsync(taskId, ct));
}
