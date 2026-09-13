namespace MicrosoftMcp.Calendar.Tests;

using MicrosoftMcp.Common;

/// <summary>In-memory fake of <see cref="IGraphCalendarService"/>. No Graph, no network.</summary>
internal sealed class FakeGraphCalendarService : IGraphCalendarService
{
    private sealed class StoredEvent
    {
        public required string Id { get; init; }
        public required string CalendarId { get; init; }
        public string Subject { get; set; } = string.Empty;
        public DateTimeOffset Start { get; init; }
        public DateTimeOffset End { get; init; }
        public bool IsAllDay { get; init; }
        public bool IsCancelled { get; init; }
        public string? Location { get; init; }
    }

    private readonly Dictionary<string, CalendarInfo> _calendars = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, StoredEvent> _events = new(StringComparer.OrdinalIgnoreCase);

    public FakeGraphCalendarService()
    {
        _calendars["cal-default"] = new CalendarInfo("cal-default", "Kalender", "auto", true, true);
        _calendars["cal-birth"] = new CalendarInfo("cal-birth", "Geburtstage", "preset1", false, false);

        var now = DateTimeOffset.UtcNow;
        var today = now.Date;
        Add(new StoredEvent
        {
            Id = "e-standup", CalendarId = "cal-default", Subject = "Daily Standup",
            Start = now.AddHours(2), End = now.AddHours(2).AddMinutes(15),
            Location = "Teams"
        });
        Add(new StoredEvent
        {
            Id = "e-holiday", CalendarId = "cal-default", Subject = "Feiertag",
            Start = today, End = today.AddDays(1), IsAllDay = true
        });
        Add(new StoredEvent
        {
            Id = "e-old", CalendarId = "cal-default", Subject = "Abgesagt",
            Start = now.AddDays(-2), End = now.AddDays(-2).AddHours(1),
            IsCancelled = true
        });
        Add(new StoredEvent
        {
            Id = "e-bday", CalendarId = "cal-birth", Subject = "Geburtstag Anna",
            Start = today.AddDays(3), End = today.AddDays(4), IsAllDay = true
        });
    }

    private void Add(StoredEvent e) => _events[e.Id] = e;

    private CalendarInfo Calendar(string? calendarId)
    {
        if (string.IsNullOrWhiteSpace(calendarId)
            || string.Equals(calendarId.Trim(), "default", StringComparison.OrdinalIgnoreCase))
        {
            return _calendars.Values.First(c => c.IsDefault);
        }

        return _calendars.TryGetValue(calendarId.Trim(), out var cal)
            ? cal
            : throw MailServiceException.CalendarNotFound(calendarId, "fake");
    }

    private static EventSummary ToSummary(StoredEvent e) => new(
        e.Id, e.Subject,
        e.Start.ToString("yyyy-MM-ddTHH:mm:ss"), "UTC",
        e.End.ToString("yyyy-MM-ddTHH:mm:ss"), "UTC",
        e.IsAllDay, e.Location, "boss@example.com", 2, "busy",
        e.IsCancelled, "preview", null);

    private static EventDetail ToDetail(StoredEvent e) => new(
        e.Id, e.Subject,
        e.Start.ToString("yyyy-MM-ddTHH:mm:ss"), "UTC",
        e.End.ToString("yyyy-MM-ddTHH:mm:ss"), "UTC",
        e.IsAllDay, e.Location, "boss@example.com",
        [new AttendeeDto("A", "a@x.y", "accepted")],
        null, null, null, true, "busy", e.IsCancelled, "preview", "body", null);

    public Task<IReadOnlyList<CalendarInfo>> ListCalendarsAsync(CancellationToken ct = default) =>
        Task.FromResult<IReadOnlyList<CalendarInfo>>([.. _calendars.Values]);

    public Task<IReadOnlyList<EventSummary>> ListEventsAsync(
        string? calendarId = null,
        string? timeMin = null,
        string? timeMax = null,
        int top = 50,
        CancellationToken ct = default)
    {
        var cal = Calendar(calendarId);
        DateTimeOffset start = ParseBound(timeMin, "timeMin") ?? DateTimeOffset.UtcNow;
        DateTimeOffset end = ParseBound(timeMax, "timeMax") ?? DateTimeOffset.UtcNow.AddDays(7);
        int take = Math.Clamp(top, 1, 200);
        return Task.FromResult<IReadOnlyList<EventSummary>>(
            [.. _events.Values
                .Where(e => e.CalendarId == cal.Id && e.Start < end && e.End > start)
                .OrderBy(e => e.Start)
                .Take(take).Select(ToSummary)]);
    }

    public Task<IReadOnlyList<EventSummary>> SearchEventsAsync(
        string query, int top = 25, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            throw MailServiceException.InvalidRequest(
                "Query must not be empty.", "pass a subject or keyword");
        }

        int take = Math.Clamp(top, 1, 50);
        return Task.FromResult<IReadOnlyList<EventSummary>>(
            [.. _events.Values
                .Where(e => e.Subject.Contains(query, StringComparison.OrdinalIgnoreCase))
                .Take(take).Select(ToSummary)]);
    }

    public Task<EventDetail> GetEventAsync(
        string eventId, string? calendarId = null, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(eventId))
        {
            throw MailServiceException.InvalidRequest(
                "eventId must not be empty.", "use an id from calendar_list_events or calendar_search_events");
        }

        return Task.FromResult(
            _events.TryGetValue(eventId.Trim(), out var e)
                ? ToDetail(e)
                : throw MailServiceException.EventNotFound(eventId, "fake"));
    }

    private static DateTimeOffset? ParseBound(string? value, string what)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return DateTimeOffset.TryParse(value, out var parsed)
            ? parsed
            : throw MailServiceException.InvalidRequest(
                $"{what} '{value}' is not a valid date/time.",
                "use ISO format, e.g. \"2026-09-14T00:00:00\"");
    }
}
