namespace MicrosoftMcp.Calendar;

/// <summary>Calendar access. No delete or move operations.</summary>
public interface IGraphCalendarService
{
    Task<IReadOnlyList<CalendarInfo>> ListCalendarsAsync(CancellationToken ct = default);
    Task<IReadOnlyList<EventSummary>> ListEventsAsync(
        string? calendarId = null,
        string? timeMin = null,
        string? timeMax = null,
        int top = 50,
        CancellationToken ct = default);
    Task<IReadOnlyList<EventSummary>> SearchEventsAsync(
        string query, int top = 25, CancellationToken ct = default);
    Task<EventDetail> GetEventAsync(
        string eventId, string? calendarId = null, CancellationToken ct = default);
    Task<EventDetail> CreateEventAsync(
        string subject,
        string start,
        string end,
        string? calendarId = null,
        string? body = null,
        string? location = null,
        IReadOnlyList<string>? attendees = null,
        bool isAllDay = false,
        int? reminderMinutesBeforeStart = null,
        CancellationToken ct = default);
    Task<EventDetail> UpdateEventAsync(
        string eventId,
        string? calendarId = null,
        string? subject = null,
        string? start = null,
        string? end = null,
        string? body = null,
        string? location = null,
        IReadOnlyList<string>? attendees = null,
        bool? isAllDay = null,
        int? reminderMinutesBeforeStart = null,
        CancellationToken ct = default);
}
