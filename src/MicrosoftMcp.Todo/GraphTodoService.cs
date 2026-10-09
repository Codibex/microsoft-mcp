using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using MicrosoftMcp.Common;
using GraphTaskStatus = Microsoft.Graph.Models.TaskStatus;

namespace MicrosoftMcp.Todo;

/// <summary>Microsoft To Do access for the signed-in user (/me/todo).
/// Add + complete only – no delete.</summary>
public sealed class GraphTodoService : IGraphTodoService
{
    // NOTE: the To Do backend (Exchange) rejects $select with
    // 400 RequestBroker--ParseUri on lists, list-tasks and single-task
    // GETs (confirmed by docs-adjacent reports + StackOverflow 74298587 /
    // 79609269), so no request below sends $select. $top and $filter on
    // status are supported; ordering is done client-side.
    // The request has no $orderby (support undocumented), so newest-first
    // ordering is applied client-side after paging. The scan stops after
    // MaxScannedTasks raw items so one call cannot page a pathological
    // list forever (same pattern as SharePoint's MaxScannedChildren).
    private const int MaxScannedTasks = 500;
    private readonly GraphServiceClient _client;

    public GraphTodoService(
        GraphServiceClient client,
        IOptions<GraphAuthOptions> options)
    {
        // To Do via /me/* requires a signed-in user. App-only would
        // need a different permission/path scheme (out of scope).
        if (options.Value.AuthMode == AuthMode.AppOnly)
        {
            throw GraphServiceException.AuthMisconfigured(
                "The To Do host requires delegated auth. Set Graph:AuthMode to Delegated.");
        }

        _client = client;
    }

    public async Task<IReadOnlyList<TodoListInfo>> ListListsAsync(CancellationToken ct = default)
    {
        var lists = new List<TodoTaskList>();
        var page = await _client.Me.Todo.Lists.GetAsync(c =>
        {
            c.QueryParameters.Top = 100;
        }, ct).ConfigureAwait(false);
        while (page is not null)
        {
            lists.AddRange(page.Value ?? []);
            if (string.IsNullOrWhiteSpace(page.OdataNextLink))
            {
                break;
            }

            page = await _client.Me.Todo.Lists
                .WithUrl(page.OdataNextLink)
                .GetAsync(cancellationToken: ct)
                .ConfigureAwait(false);
        }

        return [.. lists.Select(TodoMapper.MapList)];
    }

    public async Task<IReadOnlyList<TodoTaskSummary>> ListTasksAsync(
        string? listId = null, bool includeCompleted = false, int top = 50,
        CancellationToken ct = default)
    {
        string lid = await ResolveListIdAsync(listId, "todo_list_tasks", ct).ConfigureAwait(false);
        int take = Math.Clamp(top, 1, 100);
        var matched = new List<TodoTask>();
        int scanned = 0;
        var page = await _client.Me.Todo.Lists[lid].Tasks.GetAsync(c =>
        {
            c.QueryParameters.Top = Math.Min(take, 50);
            if (!includeCompleted)
            {
                c.QueryParameters.Filter = "status ne 'completed'";
            }
        }, ct).ConfigureAwait(false);
        // Page through (up to the scan cap) before sorting: stopping at
        // `take` matches could omit newer tasks sitting behind nextLink.
        while (page is not null && scanned < MaxScannedTasks)
        {
            IReadOnlyList<TodoTask> raw = page.Value ?? [];
            scanned += raw.Count;
            matched.AddRange(includeCompleted
                ? raw
                : raw.Where(t => t.Status != GraphTaskStatus.Completed));
            if (string.IsNullOrWhiteSpace(page.OdataNextLink) || scanned >= MaxScannedTasks)
            {
                break;
            }

            page = await _client.Me.Todo.Lists[lid].Tasks
                .WithUrl(page.OdataNextLink)
                .GetAsync(cancellationToken: ct)
                .ConfigureAwait(false);
        }

        return [.. matched
            .OrderByDescending(t => t.CreatedDateTime ?? DateTimeOffset.MinValue)
            .Take(take)
            .Select(t => TodoMapper.MapSummary(t, lid))];
    }

    public async Task<TodoTaskDetail> AddTaskAsync(
        string subject, string? listId = null, string? dueDateTime = null,
        string? importance = null, string? body = null,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(subject))
        {
            throw GraphServiceException.InvalidRequest(
                "Subject must not be empty.",
                "pass the commitment text, e.g. \"Send the offer by Friday\"");
        }

        string lid = await ResolveListIdAsync(listId, "todo_add_task", ct).ConfigureAwait(false);
        var task = new TodoTask
        {
            Title = subject.Trim()
        };
        if (!string.IsNullOrWhiteSpace(importance))
        {
            task.Importance = ParseImportance(importance);
        }

        if (!string.IsNullOrWhiteSpace(dueDateTime))
        {
            task.DueDateTime = ParseDateTime(dueDateTime, "dueDateTime");
        }

        if (!string.IsNullOrWhiteSpace(body))
        {
            task.Body = new ItemBody
            {
                Content = body.Trim(),
                ContentType = BodyType.Text
            };
        }

        TodoTask? created;
        try
        {
            created = await _client.Me.Todo.Lists[lid].Tasks
                .PostAsync(task, cancellationToken: ct).ConfigureAwait(false);
        }
        catch (ApiException ex) when (ex.ResponseStatusCode == 404)
        {
            throw GraphServiceException.TodoListNotFound(lid, "todo_add_task");
        }

        return created?.Id is null
            ? throw GraphServiceException.GraphError(200, null, "Graph returned no task after create.")
            : TodoMapper.MapDetail(created, lid);
    }

    public async Task<TodoTaskDetail> CompleteTaskAsync(
        string taskId, string? listId = null, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(taskId))
        {
            throw GraphServiceException.InvalidRequest(
                "taskId must not be empty.",
                "call todo_list_tasks to get valid task ids");
        }

        string tid = taskId.Trim();
        if (!string.IsNullOrWhiteSpace(listId))
        {
            // Validate the list first so an invalid list reports
            // [todo-list-not-found] instead of [todo-task-not-found].
            string lid = await ResolveListIdAsync(listId, "todo_complete_task", ct).ConfigureAwait(false);
            return await PatchCompletedAsync(lid, tid, ct).ConfigureAwait(false);
        }

        // No list given: discover the owning list (Graph addresses tasks
        // per list: /me/todo/lists/{listId}/tasks/{taskId}).
        IReadOnlyList<TodoListInfo> lists = await ListListsAsync(ct).ConfigureAwait(false);
        foreach (TodoListInfo list in lists)
        {
            try
            {
                TodoTask? existing = await _client.Me.Todo.Lists[list.Id].Tasks[tid]
                    .GetAsync(cancellationToken: ct).ConfigureAwait(false);
                if (existing?.Id is null)
                {
                    continue;
                }

                return await PatchCompletedAsync(list.Id, tid, ct).ConfigureAwait(false);
            }
            catch (ApiException ex) when (ex.ResponseStatusCode == 404)
            {
                continue;
            }
        }

        throw GraphServiceException.TodoTaskNotFound(tid, "todo_complete_task");
    }

    private async Task<TodoTaskDetail> PatchCompletedAsync(string listId, string taskId, CancellationToken ct)
    {
        TodoTask? patched;
        try
        {
            patched = await _client.Me.Todo.Lists[listId].Tasks[taskId]
                .PatchAsync(new TodoTask { Status = GraphTaskStatus.Completed }, cancellationToken: ct)
                .ConfigureAwait(false);
        }
        catch (ApiException ex) when (ex.ResponseStatusCode == 404)
        {
            throw GraphServiceException.TodoTaskNotFound(taskId, "todo_complete_task");
        }

        if (patched?.Id is null)
        {
            // The PATCH succeeded but returned no body: re-read for the result.
            try
            {
                patched = await _client.Me.Todo.Lists[listId].Tasks[taskId]
                    .GetAsync(cancellationToken: ct).ConfigureAwait(false);
            }
            catch (ApiException ex) when (ex.ResponseStatusCode == 404)
            {
                throw GraphServiceException.TodoTaskNotFound(taskId, "todo_complete_task");
            }
        }

        return patched?.Id is null
            ? throw GraphServiceException.TodoTaskNotFound(taskId, "todo_complete_task")
            : TodoMapper.MapDetail(patched, listId);
    }

    private async Task<string> ResolveListIdAsync(string? listId, string operation, CancellationToken ct)
    {
        if (!string.IsNullOrWhiteSpace(listId))
        {
            string lid = listId.Trim();
            try
            {
                TodoTaskList? existing = await _client.Me.Todo.Lists[lid]
                    .GetAsync(cancellationToken: ct).ConfigureAwait(false);
                if (existing?.Id is null)
                {
                    throw GraphServiceException.TodoListNotFound(lid, operation);
                }

                return existing.Id;
            }
            catch (ApiException ex) when (ex.ResponseStatusCode == 404)
            {
                throw GraphServiceException.TodoListNotFound(lid, operation);
            }
        }

        // Well-known "Tasks" fallback: the default list, else the first list.
        IReadOnlyList<TodoListInfo> lists = await ListListsAsync(ct).ConfigureAwait(false);
        TodoListInfo? @default = lists.FirstOrDefault(l =>
            string.Equals(l.WellknownListName, nameof(WellknownListName.DefaultList), StringComparison.OrdinalIgnoreCase));
        return @default?.Id
            ?? lists.FirstOrDefault()?.Id
            ?? throw GraphServiceException.InvalidRequest(
                "No To Do lists found.",
                "create a list in Microsoft To Do first, then retry");
    }

    private static Importance ParseImportance(string value)
    {
        return value.Trim().ToLowerInvariant() switch
        {
            "low" => Importance.Low,
            "normal" => Importance.Normal,
            "high" => Importance.High,
            _ => throw GraphServiceException.InvalidRequest(
                $"Unknown importance '{value}'.",
                "use low, normal or high")
        };
    }

    private static DateTimeTimeZone ParseDateTime(string value, string what)
    {
        if (!DateTimeOffset.TryParse(value.Trim(), out DateTimeOffset parsed))
        {
            throw GraphServiceException.InvalidRequest(
                $"{what} '{value}' is not a valid date/time.",
                "use ISO format, e.g. 2026-10-15T17:00:00Z");
        }

        DateTimeOffset utc = parsed.ToUniversalTime();
        return new DateTimeTimeZone
        {
            DateTime = utc.ToString("yyyy-MM-ddTHH:mm:ss.fffffff"),
            TimeZone = "UTC"
        };
    }
}
