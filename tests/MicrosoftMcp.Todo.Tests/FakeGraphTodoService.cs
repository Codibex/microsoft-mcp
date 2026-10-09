namespace MicrosoftMcp.Todo.Tests;

using MicrosoftMcp.Common;

/// <summary>In-memory fake of <see cref="IGraphTodoService"/>. No Graph, no network.</summary>
internal sealed class FakeGraphTodoService : IGraphTodoService
{
    private sealed class StoredTask
    {
        public required string Id { get; init; }
        public required string ListId { get; init; }
        public string Title { get; set; } = string.Empty;
        public string Status { get; set; } = "NotStarted";
        public string? Importance { get; set; }
        public string? Due { get; set; }
        public string? Body { get; set; }
    }

    private readonly Dictionary<string, TodoListInfo> _lists = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, StoredTask> _tasks = new(StringComparer.OrdinalIgnoreCase);

    public FakeGraphTodoService()
    {
        _lists["l-tasks"] = new TodoListInfo("l-tasks", "Tasks", true, false, "DefaultList");
        _lists["l-shared"] = new TodoListInfo("l-shared", "Family", false, true, "None");
        Add(new StoredTask { Id = "t-1", ListId = "l-tasks", Title = "Offer senden", Status = "NotStarted", Importance = "High" });
        Add(new StoredTask { Id = "t-done", ListId = "l-tasks", Title = "Erledigt", Status = "Completed" });
        Add(new StoredTask { Id = "t-shared", ListId = "l-shared", Title = "Einkaufen", Status = "NotStarted" });
    }

    private void Add(StoredTask t) => _tasks[t.Id] = t;

    private static void Require(string? value, string what, string hint)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw GraphServiceException.InvalidRequest($"{what} must not be empty.", hint);
        }
    }

    private string ResolveList(string? listId)
    {
        if (!string.IsNullOrWhiteSpace(listId))
        {
            string lid = listId.Trim();
            if (!_lists.ContainsKey(lid))
            {
                throw GraphServiceException.TodoListNotFound(lid, "fake");
            }

            return lid;
        }

        return _lists.Values.FirstOrDefault(l => l.WellknownListName == "DefaultList")?.Id
            ?? _lists.Values.First().Id;
    }

    private static TodoTaskSummary ToSummary(StoredTask t) => new(
        t.Id, t.ListId, t.Title, t.Status, t.Importance, t.Due,
        DateTimeOffset.UtcNow, DateTimeOffset.UtcNow, false);

    private static TodoTaskDetail ToDetail(StoredTask t) => new(
        t.Id, t.ListId, t.Title, t.Status, t.Importance, t.Due, null, null,
        DateTimeOffset.UtcNow, DateTimeOffset.UtcNow, t.Body, t.Body is null ? null : "Text", []);

    public Task<IReadOnlyList<TodoListInfo>> ListListsAsync(CancellationToken ct = default) =>
        Task.FromResult<IReadOnlyList<TodoListInfo>>([.. _lists.Values]);

    public Task<IReadOnlyList<TodoTaskSummary>> ListTasksAsync(
        string? listId = null, bool includeCompleted = false, int top = 50,
        CancellationToken ct = default)
    {
        string lid = ResolveList(listId);
        var tasks = _tasks.Values
            .Where(t => t.ListId == lid)
            .Where(t => includeCompleted || t.Status != "Completed")
            .Take(Math.Clamp(top, 1, 100))
            .Select(ToSummary)
            .ToList();
        return Task.FromResult<IReadOnlyList<TodoTaskSummary>>(tasks);
    }

    public Task<TodoTaskDetail> AddTaskAsync(
        string subject, string? listId = null, string? dueDateTime = null,
        string? importance = null, string? body = null,
        CancellationToken ct = default)
    {
        Require(subject, "subject", "pass the commitment text");
        if (!string.IsNullOrWhiteSpace(importance)
            && !new[] { "low", "normal", "high" }.Contains(importance.Trim().ToLowerInvariant()))
        {
            throw GraphServiceException.InvalidRequest(
                $"Unknown importance '{importance}'.", "use low, normal or high");
        }

        if (!string.IsNullOrWhiteSpace(dueDateTime)
            && !DateTimeOffset.TryParse(dueDateTime.Trim(), out _))
        {
            throw GraphServiceException.InvalidRequest(
                $"dueDateTime '{dueDateTime}' is not a valid date/time.", "use ISO format");
        }

        string lid = ResolveList(listId);
        var task = new StoredTask
        {
            Id = $"t-{Guid.NewGuid():N}",
            ListId = lid,
            Title = subject.Trim(),
            Importance = importance?.Trim(),
            Due = dueDateTime?.Trim(),
            Body = body?.Trim()
        };
        Add(task);
        return Task.FromResult(ToDetail(task));
    }

    public Task<TodoTaskDetail> CompleteTaskAsync(
        string taskId, string? listId = null, CancellationToken ct = default)
    {
        Require(taskId, "taskId", "call todo_list_tasks to get valid task ids");
        string tid = taskId.Trim();
        if (!string.IsNullOrWhiteSpace(listId))
        {
            string lid = listId.Trim();
            if (!_lists.ContainsKey(lid))
            {
                throw GraphServiceException.TodoListNotFound(lid, "fake");
            }

            if (!_tasks.TryGetValue(tid, out var task) || task.ListId != lid)
            {
                throw GraphServiceException.TodoTaskNotFound(tid, "fake");
            }

            task.Status = "Completed";
            return Task.FromResult(ToDetail(task));
        }

        if (!_tasks.TryGetValue(tid, out var found))
        {
            throw GraphServiceException.TodoTaskNotFound(tid, "fake");
        }

        found.Status = "Completed";
        return Task.FromResult(ToDetail(found));
    }
}
