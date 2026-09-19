using System.ComponentModel;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using MicrosoftMcp.Common;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace MicrosoftMcp.Calendar;

public static class CalendarServiceRegistration
{
    /// <summary>Compatibility adapter for callers using the legacy shared policy.
    /// Legacy recipient restrictions apply to calendar attendees.</summary>
    public static IServiceCollection AddCalendar(
        this IServiceCollection services, MessagingPolicyOptions legacyPolicy)
    {
        ArgumentNullException.ThrowIfNull(legacyPolicy);
        return services.AddCalendar(new CalendarPolicyOptions
        {
            RequireInternalAttendees = legacyPolicy.RequireInternalRecipients,
            AllowedAttendeeDomains = [.. legacyPolicy.AllowedRecipientDomains],
            AllowedAttendeeAddresses = [.. legacyPolicy.AllowedRecipientAddresses]
        });
    }

    public static IServiceCollection AddCalendar(
        this IServiceCollection services, CalendarPolicyOptions? policy = null)
    {
        services.AddSingleton<IGraphCalendarService, GraphCalendarService>();
        services.AddSingleton<IOptions<CalendarPolicyOptions>>(
            Options.Create(policy ?? new CalendarPolicyOptions()));
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

    [McpServerTool, Description("Create an event. Requires delegated Calendars.ReadWrite permission; no invitation is sent without attendees.")]
    public Task<CallToolResult> calendar_create_event(
        [Description("Event subject/title")] string subject,
        [Description("Start in ISO format, e.g. 2026-09-14T14:00:00Z")] string start,
        [Description("End in ISO format, e.g. 2026-09-14T15:00:00Z")] string end,
        [Description("Calendar id or \"default\" (optional)")] string? calendarId = null,
        [Description("Plain-text event body (optional)")] string? body = null,
        [Description("Location name (optional)")] string? location = null,
        [Description("Attendee SMTP addresses (optional)")] IReadOnlyList<string>? attendees = null,
        [Description("Whether the event spans whole days")] bool isAllDay = false,
        [Description("Minutes before start for the reminder (optional)")] int? reminderMinutesBeforeStart = null,
        CancellationToken ct = default) =>
        InvokeAsync("calendar_create_event", () => calendar.CreateEventAsync(
            subject, start, end, calendarId, body, location, attendees,
            isAllDay, reminderMinutesBeforeStart, ct));

    [McpServerTool, Description("Update supplied fields of an event. Requires delegated Calendars.ReadWrite permission; no delete or move is offered.")]
    public Task<CallToolResult> calendar_update_event(
        [Description("Graph event id")] string eventId,
        [Description("Calendar id or \"default\" (optional)")] string? calendarId = null,
        [Description("New subject/title (optional)")] string? subject = null,
        [Description("New start in ISO format (optional)")] string? start = null,
        [Description("New end in ISO format (optional)")] string? end = null,
        [Description("New plain-text body (optional)")] string? body = null,
        [Description("New location name (optional)")] string? location = null,
        [Description("Replacement attendee SMTP addresses (optional; empty clears)")] IReadOnlyList<string>? attendees = null,
        [Description("Whether the event spans whole days (optional)")] bool? isAllDay = null,
        [Description("New minutes before start for the reminder (optional)")] int? reminderMinutesBeforeStart = null,
        CancellationToken ct = default) =>
        InvokeAsync("calendar_update_event", () => calendar.UpdateEventAsync(
            eventId, calendarId, subject, start, end, body, location, attendees,
            isAllDay, reminderMinutesBeforeStart, ct));
}
