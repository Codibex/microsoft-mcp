using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using MicrosoftMcp.Common;

namespace MicrosoftMcp.Calendar;

/// <summary>Read-only calendar access on the default calendar or a specific
/// calendar id. No write operations exist on this service by design.</summary>
public sealed class GraphCalendarService(
    GraphServiceClient client,
    IOptions<GraphAuthOptions> options) : IGraphCalendarService
{
    private static readonly string[] EventSelect =
        ["id", "subject", "bodyPreview", "body", "start", "end", "isAllDay",
         "location", "organizer", "attendees", "recurrence", "showAs",
         "sensitivity", "isCancelled", "isOrganizer", "onlineMeeting",
         "webLink", "responseStatus", "seriesMasterId", "type"];

    private readonly GraphAuthOptions _options = options.Value;
    private bool IsMe => string.Equals(_options.UserIdOrUpn, "me", StringComparison.OrdinalIgnoreCase);

    public async Task<IReadOnlyList<CalendarInfo>> ListCalendarsAsync(CancellationToken ct = default)
    {
        if (IsMe)
        {
            var page = await client.Me.Calendars.GetAsync(c =>
                c.QueryParameters.Select = ["id", "name", "color", "isDefaultCalendar", "canEdit"], ct)
                .ConfigureAwait(false);
            return [.. (page?.Value ?? []).Select(CalendarMapper.MapCalendar)];
        }

        var userPage = await client.Users[_options.UserIdOrUpn].Calendars.GetAsync(c =>
            c.QueryParameters.Select = ["id", "name", "color", "isDefaultCalendar", "canEdit"], ct)
            .ConfigureAwait(false);
        return [.. (userPage?.Value ?? []).Select(CalendarMapper.MapCalendar)];
    }

    public async Task<IReadOnlyList<EventSummary>> ListEventsAsync(
        string? calendarId = null,
        string? timeMin = null,
        string? timeMax = null,
        int top = 50,
        CancellationToken ct = default)
    {
        string start = RequireWindowBound(timeMin, "timeMin")
            ?? DateTimeOffset.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss");
        string end = RequireWindowBound(timeMax, "timeMax")
            ?? DateTimeOffset.UtcNow.AddDays(7).ToString("yyyy-MM-ddTHH:mm:ss");
        int take = Math.Clamp(top, 1, 200);

        if (IsDefaultCalendar(calendarId))
        {
            if (IsMe)
            {
                var page = await client.Me.Calendar.CalendarView.GetAsync(c =>
                {
                    c.QueryParameters.StartDateTime = start;
                    c.QueryParameters.EndDateTime = end;
                    c.QueryParameters.Top = take;
                    c.QueryParameters.Select = EventSelect;
                    c.QueryParameters.Orderby = ["start/dateTime"];
                }, ct).ConfigureAwait(false);
                return [.. (page?.Value ?? []).Select(CalendarMapper.MapSummary)];
            }

            var userPage = await client.Users[_options.UserIdOrUpn].Calendar.CalendarView.GetAsync(c =>
            {
                c.QueryParameters.StartDateTime = start;
                c.QueryParameters.EndDateTime = end;
                c.QueryParameters.Top = take;
                c.QueryParameters.Select = EventSelect;
                c.QueryParameters.Orderby = ["start/dateTime"];
            }, ct).ConfigureAwait(false);
            return [.. (userPage?.Value ?? []).Select(CalendarMapper.MapSummary)];
        }

        if (IsMe)
        {
            var page = await client.Me.Calendars[calendarId!].CalendarView.GetAsync(c =>
            {
                c.QueryParameters.StartDateTime = start;
                c.QueryParameters.EndDateTime = end;
                c.QueryParameters.Top = take;
                c.QueryParameters.Select = EventSelect;
                c.QueryParameters.Orderby = ["start/dateTime"];
            }, ct).ConfigureAwait(false);
            return [.. (page?.Value ?? []).Select(CalendarMapper.MapSummary)];
        }

        var specific = await client.Users[_options.UserIdOrUpn].Calendars[calendarId!].CalendarView.GetAsync(c =>
        {
            c.QueryParameters.StartDateTime = start;
            c.QueryParameters.EndDateTime = end;
            c.QueryParameters.Top = take;
            c.QueryParameters.Select = EventSelect;
            c.QueryParameters.Orderby = ["start/dateTime"];
        }, ct).ConfigureAwait(false);
        return [.. (specific?.Value ?? []).Select(CalendarMapper.MapSummary)];
    }

    public async Task<IReadOnlyList<EventSummary>> SearchEventsAsync(
        string query, int top = 25, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            throw GraphServiceException.InvalidRequest(
                "Query must not be empty.",
                "pass a subject or keyword, e.g. \"dentist\"");
        }

        int take = Math.Clamp(top, 1, 50);
        if (IsMe)
        {
            var page = await client.Me.Events.GetAsync(c =>
            {
                c.QueryParameters.Search = $"\"{query.Trim()}\"";
                c.QueryParameters.Top = take;
                c.QueryParameters.Select = EventSelect;
            }, ct).ConfigureAwait(false);
            return [.. (page?.Value ?? []).Select(CalendarMapper.MapSummary)];
        }

        var userPage = await client.Users[_options.UserIdOrUpn].Events.GetAsync(c =>
        {
            c.QueryParameters.Search = $"\"{query.Trim()}\"";
            c.QueryParameters.Top = take;
            c.QueryParameters.Select = EventSelect;
        }, ct).ConfigureAwait(false);
        return [.. (userPage?.Value ?? []).Select(CalendarMapper.MapSummary)];
    }

    public async Task<EventDetail> GetEventAsync(
        string eventId, string? calendarId = null, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(eventId))
        {
            throw GraphServiceException.InvalidRequest(
                "eventId must not be empty.",
                "use an id from calendar_list_events or calendar_search_events");
        }

        Event? ev = (IsDefaultCalendar(calendarId), IsMe) switch
        {
            (true, true) => await client.Me.Calendar.Events[eventId.Trim()].GetAsync(c =>
                c.QueryParameters.Select = EventSelect, ct).ConfigureAwait(false),
            (true, false) => await client.Users[_options.UserIdOrUpn].Calendar.Events[eventId.Trim()].GetAsync(c =>
                c.QueryParameters.Select = EventSelect, ct).ConfigureAwait(false),
            (false, true) => await client.Me.Calendars[calendarId!].Events[eventId.Trim()].GetAsync(c =>
                c.QueryParameters.Select = EventSelect, ct).ConfigureAwait(false),
            (false, false) => await client.Users[_options.UserIdOrUpn].Calendars[calendarId!].Events[eventId.Trim()].GetAsync(c =>
                c.QueryParameters.Select = EventSelect, ct).ConfigureAwait(false)
        };

        return ev is null
            ? throw GraphServiceException.EventNotFound(eventId, "calendar_read_event")
            : CalendarMapper.MapDetail(ev);
    }

    private static bool IsDefaultCalendar(string? calendarId) =>
        string.IsNullOrWhiteSpace(calendarId)
        || string.Equals(calendarId.Trim(), "default", StringComparison.OrdinalIgnoreCase);

    private static string? RequireWindowBound(string? value, string what)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        if (!DateTimeOffset.TryParse(value, out _))
        {
            throw GraphServiceException.InvalidRequest(
                $"{what} '{value}' is not a valid date/time.",
                "use ISO format, e.g. \"2026-09-14T00:00:00\"");
        }

        return value.Trim();
    }
}
