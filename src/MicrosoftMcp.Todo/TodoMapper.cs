using Microsoft.Graph.Models;

namespace MicrosoftMcp.Todo;

internal static class TodoMapper
{
    internal static TodoListInfo MapList(TodoTaskList list) => new(
        list.Id ?? string.Empty,
        list.DisplayName ?? string.Empty,
        list.IsOwner,
        list.IsShared,
        list.WellknownListName?.ToString());

    internal static TodoTaskSummary MapSummary(TodoTask task, string listId) => new(
        task.Id ?? string.Empty,
        listId,
        task.Title,
        task.Status?.ToString(),
        task.Importance?.ToString(),
        FormatDt(task.DueDateTime),
        ToOffset(task.CreatedDateTime),
        ToOffset(task.LastModifiedDateTime),
        task.HasAttachments);

    internal static TodoTaskDetail MapDetail(TodoTask task, string listId) => new(
        task.Id ?? string.Empty,
        listId,
        task.Title,
        task.Status?.ToString(),
        task.Importance?.ToString(),
        FormatDt(task.DueDateTime),
        FormatDt(task.ReminderDateTime),
        FormatDt(task.CompletedDateTime),
        ToOffset(task.CreatedDateTime),
        ToOffset(task.LastModifiedDateTime),
        task.Body?.Content?.Trim(),
        task.Body?.ContentType?.ToString(),
        task.Categories ?? []);

    internal static string? FormatDt(DateTimeTimeZone? dt)
    {
        if (dt?.DateTime is null)
        {
            return null;
        }

        return string.IsNullOrWhiteSpace(dt.TimeZone)
            ? dt.DateTime
            : $"{dt.DateTime} ({dt.TimeZone})";
    }

    internal static DateTimeOffset? ToOffset(DateTimeOffset? value) => value;
}
