namespace MicrosoftMcp.Todo;

/// <summary>Microsoft To Do access for the signed-in user (/me/todo).
/// Add + complete only – no delete (same spirit as the mail drafts-only rule).</summary>
public interface IGraphTodoService
{
    Task<IReadOnlyList<TodoListInfo>> ListListsAsync(CancellationToken ct = default);
    Task<IReadOnlyList<TodoTaskSummary>> ListTasksAsync(
        string? listId = null, bool includeCompleted = false, int top = 50,
        CancellationToken ct = default);
    Task<TodoTaskDetail> AddTaskAsync(
        string subject, string? listId = null, string? dueDateTime = null,
        string? importance = null, string? body = null,
        CancellationToken ct = default);
    Task<TodoTaskDetail> CompleteTaskAsync(
        string taskId, string? listId = null, CancellationToken ct = default);
}
