namespace MicrosoftMcp.Calendar;

/// <summary>Read-only calendar access. No create/update/delete/move.</summary>
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
}
