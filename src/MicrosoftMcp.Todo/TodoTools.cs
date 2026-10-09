using System.ComponentModel;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using MicrosoftMcp.Common;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace MicrosoftMcp.Todo;

public static class TodoServiceRegistration
{
    public static IServiceCollection AddTodo(this IServiceCollection services)
    {
        services.AddSingleton<IGraphTodoService, GraphTodoService>();
        return services;
    }
}

/// <summary>Microsoft To Do: harvest commitments from mail/meetings into
/// tracked tasks. Add + complete only – no delete is offered.</summary>
[McpServerToolType]
public sealed class TodoTools(IGraphTodoService todo, ILogger<TodoTools> log)
{
    private async Task<CallToolResult> InvokeAsync<T>(
        string operation, Func<Task<T>> call, string resource = "todo-task")
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

    [McpServerTool, Description("List To Do lists (id, name, shared flag). Needed to pick a list id. Read-only.")]
    public Task<CallToolResult> todo_list_lists(CancellationToken ct = default) =>
        InvokeAsync("todo_list_lists", () => todo.ListListsAsync(ct), "todo-list");

    [McpServerTool, Description("List tasks of a To Do list (default: well-known Tasks list). Completed tasks are hidden unless includeCompleted is true. Read-only.")]
    public Task<CallToolResult> todo_list_tasks(
        [Description("List id from todo_list_lists (optional, default Tasks list)")] string? listId = null,
        [Description("Include completed tasks")] bool includeCompleted = false,
        [Description("Max tasks 1-100")] int top = 50,
        CancellationToken ct = default) =>
        InvokeAsync("todo_list_tasks", () => todo.ListTasksAsync(listId, includeCompleted, top, ct));

    [McpServerTool, Description("Add a To Do task (visible in the user's To Do app, syncs to the phone). No delete offered.")]
    public Task<CallToolResult> todo_add_task(
        [Description("Task title, e.g. \"Send the offer by Friday\"")] string subject,
        [Description("List id from todo_list_lists (optional, default Tasks list)")] string? listId = null,
        [Description("Due date/time ISO, e.g. 2026-10-15T17:00:00Z (optional)")] string? dueDateTime = null,
        [Description("Importance low|normal|high (optional)")] string? importance = null,
        [Description("Plain-text notes (optional)")] string? body = null,
        CancellationToken ct = default) =>
        InvokeAsync("todo_add_task", () => todo.AddTaskAsync(subject, listId, dueDateTime, importance, body, ct));

    [McpServerTool, Description("Mark a To Do task as completed (the one mutation an assistant needs after doing the work).")]
    public Task<CallToolResult> todo_complete_task(
        [Description("Task id from todo_list_tasks")] string taskId,
        [Description("List id from todo_list_lists (optional, speeds up completion)")] string? listId = null,
        CancellationToken ct = default) =>
        InvokeAsync("todo_complete_task", () => todo.CompleteTaskAsync(taskId, listId, ct));
}
