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
        public DateTimeOffset Start { get; set; }
        public DateTimeOffset End { get; set; }
        public bool IsAllDay { get; set; }
        public bool IsCancelled { get; init; }
        public string? Location { get; set; }
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
            : throw GraphServiceException.CalendarNotFound(calendarId, "fake");
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
            throw GraphServiceException.InvalidRequest(
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
            throw GraphServiceException.InvalidRequest(
                "eventId must not be empty.", "use an id from calendar_list_events or calendar_search_events");
        }

        return Task.FromResult(
            _events.TryGetValue(eventId.Trim(), out var e)
                ? ToDetail(e)
                : throw GraphServiceException.EventNotFound(eventId, "fake"));
    }

    public Task<EventDetail> CreateEventAsync(
        string subject,
        string start,
        string end,
        string? calendarId = null,
        string? body = null,
        string? location = null,
        IReadOnlyList<string>? attendees = null,
        bool isAllDay = false,
        int? reminderMinutesBeforeStart = null,
        CancellationToken ct = default)
    {
        var calendar = Calendar(calendarId);
        if (string.IsNullOrWhiteSpace(subject))
        {
            throw GraphServiceException.InvalidRequest("Subject must not be empty.", "pass a title for the event");
        }

        DateTimeOffset startValue = ParseRequired(start, "start");
        DateTimeOffset endValue = ParseRequired(end, "end");
        if (endValue <= startValue)
        {
            throw GraphServiceException.InvalidRequest(
                "Event end must be after event start.", "pass an end time later than the start time");
        }

        ValidateReminder(reminderMinutesBeforeStart);
        var created = new StoredEvent
        {
            Id = $"e-created-{_events.Count}",
            CalendarId = calendar.Id,
            Subject = subject.Trim(),
            Start = startValue,
            End = endValue,
            IsAllDay = isAllDay,
            Location = location
        };
        Add(created);
        return Task.FromResult(ToDetail(created));
    }

    public Task<EventDetail> UpdateEventAsync(
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
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(eventId))
        {
            throw GraphServiceException.InvalidRequest("eventId must not be empty.", "use an event id");
        }

        var calendar = Calendar(calendarId);
        if (!_events.TryGetValue(eventId.Trim(), out var existing) || existing.CalendarId != calendar.Id)
        {
            throw GraphServiceException.EventNotFound(eventId, "fake");
        }

        if (subject is null && start is null && end is null && body is null
            && location is null && attendees is null && isAllDay is null
            && reminderMinutesBeforeStart is null)
        {
            throw GraphServiceException.InvalidRequest(
                "At least one event field must be supplied.", "pass a field to update");
        }

        if (subject is not null)
        {
            if (string.IsNullOrWhiteSpace(subject))
            {
                throw GraphServiceException.InvalidRequest("Subject must not be empty.", "pass a title for the event");
            }

            existing.Subject = subject.Trim();
        }

        DateTimeOffset newStart = start is null ? existing.Start : ParseRequired(start, "start");
        DateTimeOffset newEnd = end is null ? existing.End : ParseRequired(end, "end");
        if (end is not null && newEnd <= newStart || start is not null && newEnd <= newStart)
        {
            throw GraphServiceException.InvalidRequest(
                "Event end must be after event start.", "pass an end time later than the start time");
        }

        existing.Start = newStart;
        existing.End = newEnd;
        if (location is not null)
        {
            existing.Location = location;
        }

        if (isAllDay is not null)
        {
            existing.IsAllDay = isAllDay.Value;
        }

        ValidateReminder(reminderMinutesBeforeStart);
        return Task.FromResult(ToDetail(existing));
    }

    private static DateTimeOffset? ParseBound(string? value, string what)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return DateTimeOffset.TryParse(value, out var parsed)
            ? parsed
            : throw GraphServiceException.InvalidRequest(
                $"{what} '{value}' is not a valid date/time.",
                "use ISO format, e.g. \"2026-09-14T00:00:00\"");
    }

    private static DateTimeOffset ParseRequired(string value, string what) =>
        ParseBound(value, what) ?? throw GraphServiceException.InvalidRequest(
            $"{what} must not be empty.", "use an ISO date/time");

    private static void ValidateReminder(int? reminderMinutesBeforeStart)
    {
        if (reminderMinutesBeforeStart < 0)
        {
            throw GraphServiceException.InvalidRequest(
                "reminderMinutesBeforeStart must not be negative.",
                "pass zero or a positive number of minutes");
        }
    }
}
