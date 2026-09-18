using System.ComponentModel;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using MicrosoftMcp.Common;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace MicrosoftMcp.Calendar;

public static class CalendarServiceRegistration
{
    public static IServiceCollection AddCalendar(this IServiceCollection services)
    {
        services.AddSingleton<IGraphCalendarService, GraphCalendarService>();
        return services;
    }
}

[McpServerToolType]
public sealed class CalendarTools(IGraphCalendarService calendar, ILogger<CalendarTools> log)
{
    private async Task<CallToolResult> InvokeAsync<T>(
        string operation, Func<Task<T>> call, string resource = "event")
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

    [McpServerTool, Description("List calendars (id, name, default flag). Needed to pick a calendar id.")]
    public Task<CallToolResult> calendar_list_calendars(CancellationToken ct = default) =>
        InvokeAsync("calendar_list_calendars", () => calendar.ListCalendarsAsync(ct), "calendar");

    [McpServerTool, Description("List events in a time window (default: default calendar, now plus 7 days). Times in ISO format, e.g. 2026-09-14T00:00:00.")]
    public Task<CallToolResult> calendar_list_events(
        [Description("Calendar id or \"default\" (optional)")] string? calendarId = null,
        [Description("Window start ISO (optional, default now)")] string? timeMin = null,
        [Description("Window end ISO (optional, default now plus 7 days)")] string? timeMax = null,
        [Description("Max events 1-200")] int top = 50,
        CancellationToken ct = default) =>
        InvokeAsync("calendar_list_events", () => calendar.ListEventsAsync(calendarId, timeMin, timeMax, top, ct));

    [McpServerTool, Description("Search events by subject/keyword across calendars.")]
    public Task<CallToolResult> calendar_search_events(
        [Description("Subject or keyword, e.g. \"dentist\"")] string query,
        [Description("Max results 1-50")] int top = 25,
        CancellationToken ct = default) =>
        InvokeAsync("calendar_search_events", () => calendar.SearchEventsAsync(query, top, ct));

    [McpServerTool, Description("Read a full event by id (attendees, recurrence, body, online meeting link).")]
    public Task<CallToolResult> calendar_read_event(
        [Description("Graph event id")] string eventId,
        [Description("Calendar id or \"default\" (optional)")] string? calendarId = null,
        CancellationToken ct = default) =>
        InvokeAsync("calendar_read_event", () => calendar.GetEventAsync(eventId, calendarId, ct));
}
