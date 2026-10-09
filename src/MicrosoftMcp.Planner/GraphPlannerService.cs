using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using MicrosoftMcp.Common;

namespace MicrosoftMcp.Planner;

/// <summary>Read-only Planner access for the signed-in user.</summary>
public sealed class GraphPlannerService : IGraphPlannerService
{
    private static readonly string[] PlanSelect = ["id", "title", "owner", "createdDateTime"];
    private static readonly string[] TaskSelect =
        ["id", "planId", "bucketId", "title", "percentComplete", "priority",
            "dueDateTime", "startDateTime", "createdDateTime", "assignments", "hasDescription"];

    private readonly GraphServiceClient _client;

    public GraphPlannerService(
        GraphServiceClient client,
        IOptions<GraphAuthOptions> options)
    {
        // Planner via /me/* and /groups/* requires a signed-in user. App-only
        // would need a different permission/path scheme (out of scope).
        if (options.Value.AuthMode == AuthMode.AppOnly)
        {
            throw GraphServiceException.AuthMisconfigured(
                "The Planner host requires delegated auth. Set Graph:AuthMode to Delegated.");
        }

        _client = client;
    }

    public async Task<IReadOnlyList<PlannerTaskSummary>> ListMyTasksAsync(
        int top = 50, CancellationToken ct = default)
    {
        int take = Math.Clamp(top, 1, 100);
        var tasks = new List<PlannerTask>();
        var page = await _client.Me.Planner.Tasks.GetAsync(c =>
        {
            c.QueryParameters.Top = Math.Min(take, 50);
            c.QueryParameters.Select = TaskSelect;
        }, ct).ConfigureAwait(false);
        while (page is not null && tasks.Count < take)
        {
            tasks.AddRange(page.Value ?? []);
            if (string.IsNullOrWhiteSpace(page.OdataNextLink) || tasks.Count >= take)
            {
                break;
            }

            page = await _client.Me.Planner.Tasks
                .WithUrl(page.OdataNextLink)
                .GetAsync(cancellationToken: ct)
                .ConfigureAwait(false);
        }

        return [.. tasks.Take(take).Select(PlannerMapper.MapSummary)];
    }

    public async Task<IReadOnlyList<PlannerPlanInfo>> ListPlansAsync(
        string groupId, int top = 50, CancellationToken ct = default)
    {
        RequireId(groupId, "groupId", "use the Microsoft 365 group id backing the plan");
        string gid = groupId.Trim();
        int take = Math.Clamp(top, 1, 100);
        try
        {
            var page = await _client.Groups[gid].Planner.Plans.GetAsync(c =>
            {
                c.QueryParameters.Top = Math.Min(take, 50);
                c.QueryParameters.Select = PlanSelect;
            }, ct).ConfigureAwait(false);
            return [.. (page?.Value ?? []).Take(take).Select(PlannerMapper.MapPlan)];
        }
        catch (ApiException ex) when (ex.ResponseStatusCode == 404)
        {
            throw GraphServiceException.GroupNotFound(gid, "planner_list_plans");
        }
    }

    public async Task<IReadOnlyList<PlannerTaskSummary>> ListPlanTasksAsync(
        string planId, int top = 50, CancellationToken ct = default)
    {
        RequireId(planId, "planId", "call planner_list_plans with the group id to get valid plan ids");
        string pid = planId.Trim();
        int take = Math.Clamp(top, 1, 100);
        var tasks = new List<PlannerTask>();
        try
        {
            var page = await _client.Planner.Plans[pid].Tasks.GetAsync(c =>
            {
                c.QueryParameters.Top = Math.Min(take, 50);
                c.QueryParameters.Select = TaskSelect;
            }, ct).ConfigureAwait(false);
            while (page is not null && tasks.Count < take)
            {
                tasks.AddRange(page.Value ?? []);
                if (string.IsNullOrWhiteSpace(page.OdataNextLink) || tasks.Count >= take)
                {
                    break;
                }

                page = await _client.Planner.Plans[pid].Tasks
                    .WithUrl(page.OdataNextLink)
                    .GetAsync(cancellationToken: ct)
                    .ConfigureAwait(false);
            }
        }
        catch (ApiException ex) when (ex.ResponseStatusCode == 404)
        {
            throw GraphServiceException.PlannerPlanNotFound(pid, "planner_list_plan_tasks");
        }

        return [.. tasks.Take(take).Select(PlannerMapper.MapSummary)];
    }

    public async Task<PlannerTaskDetail> ReadTaskAsync(string taskId, CancellationToken ct = default)
    {
        RequireId(taskId, "taskId", "call planner_list_my_tasks or planner_list_plan_tasks to get valid task ids");
        string tid = taskId.Trim();
        PlannerTask? task;
        try
        {
            task = await _client.Planner.Tasks[tid]
                .GetAsync(c => c.QueryParameters.Select = TaskSelect, ct).ConfigureAwait(false);
        }
        catch (ApiException ex) when (ex.ResponseStatusCode == 404)
        {
            throw GraphServiceException.PlannerTaskNotFound(tid, "planner_read_task");
        }

        if (task?.Id is null)
        {
            throw GraphServiceException.PlannerTaskNotFound(tid, "planner_read_task");
        }

        PlannerTaskDetails? details = null;
        try
        {
            details = await _client.Planner.Tasks[tid].Details
                .GetAsync(cancellationToken: ct).ConfigureAwait(false);
        }
        catch (ApiException ex) when (ex.ResponseStatusCode == 404)
        {
            details = null;
        }

        string? bucketName = null;
        if (!string.IsNullOrWhiteSpace(task.BucketId))
        {
            try
            {
                PlannerBucket? bucket = await _client.Planner.Buckets[task.BucketId]
                    .GetAsync(c => c.QueryParameters.Select = ["id", "name"], ct).ConfigureAwait(false);
                bucketName = bucket?.Name;
            }
            catch (ApiException ex) when (ex.ResponseStatusCode is 403 or 404)
            {
                bucketName = null;
            }
        }

        return PlannerMapper.MapDetail(task, details, bucketName);
    }

    private static void RequireId(string value, string what, string hint)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw GraphServiceException.InvalidRequest($"{what} must not be empty.", hint);
        }
    }
}
