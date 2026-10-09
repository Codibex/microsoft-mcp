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
    private static readonly string[] ListSelect = ["id", "displayName", "isOwner", "isShared", "wellknownListName"];
    private static readonly string[] TaskSelect =
        ["id", "title", "status", "importance", "dueDateTime", "createdDateTime", "lastModifiedDateTime", "hasAttachments"];
    private static readonly string[] TaskDetailSelect =
        ["id", "title", "status", "importance", "dueDateTime", "reminderDateTime", "completedDateTime",
            "createdDateTime", "lastModifiedDateTime", "hasAttachments", "categories", "body"];

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
            c.QueryParameters.Select = ListSelect;
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
        string lid = await ResolveListIdAsync(listId, ct).ConfigureAwait(false);
        int take = Math.Clamp(top, 1, 100);
        var tasks = new List<TodoTask>();
        var page = await _client.Me.Todo.Lists[lid].Tasks.GetAsync(c =>
        {
            c.QueryParameters.Top = Math.Min(take, 50);
            c.QueryParameters.Select = TaskSelect;
            c.QueryParameters.Orderby = ["createdDateTime desc"];
            if (!includeCompleted)
            {
                c.QueryParameters.Filter = "status ne 'completed'";
            }
        }, ct).ConfigureAwait(false);
        while (page is not null && tasks.Count < take)
        {
            tasks.AddRange(page.Value ?? []);
            if (string.IsNullOrWhiteSpace(page.OdataNextLink) || tasks.Count >= take)
            {
                break;
            }

            page = await _client.Me.Todo.Lists[lid].Tasks
                .WithUrl(page.OdataNextLink)
                .GetAsync(cancellationToken: ct)
                .ConfigureAwait(false);
        }

        IEnumerable<TodoTask> filtered = includeCompleted
            ? tasks
            : tasks.Where(t => t.Status != GraphTaskStatus.Completed);
        return [.. filtered.Take(take).Select(t => TodoMapper.MapSummary(t, lid))];
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

        string lid = await ResolveListIdAsync(listId, ct).ConfigureAwait(false);
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
            string lid = listId.Trim();
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
                    .GetAsync(c => c.QueryParameters.Select = ["id"], ct).ConfigureAwait(false);
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
                    .GetAsync(c => c.QueryParameters.Select = TaskDetailSelect, ct).ConfigureAwait(false);
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

    private async Task<string> ResolveListIdAsync(string? listId, CancellationToken ct)
    {
        if (!string.IsNullOrWhiteSpace(listId))
        {
            string lid = listId.Trim();
            try
            {
                TodoTaskList? existing = await _client.Me.Todo.Lists[lid]
                    .GetAsync(c => c.QueryParameters.Select = ListSelect, ct).ConfigureAwait(false);
                if (existing?.Id is null)
                {
                    throw GraphServiceException.TodoListNotFound(lid, "todo_list_tasks");
                }

                return existing.Id;
            }
            catch (ApiException ex) when (ex.ResponseStatusCode == 404)
            {
                throw GraphServiceException.TodoListNotFound(lid, "todo_list_tasks");
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
