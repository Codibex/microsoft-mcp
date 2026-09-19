using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using MicrosoftMcp.Common;

namespace MicrosoftMcp.Calendar;

/// <summary>Calendar access on the default calendar or a specific calendar id.
/// Delete and move operations are intentionally not exposed.</summary>
public sealed class GraphCalendarService(
    GraphServiceClient client,
    IOptions<GraphAuthOptions> options,
    IOptions<CalendarPolicyOptions> policy) : IGraphCalendarService
{
    private static readonly string[] EventSelect =
        ["id", "subject", "bodyPreview", "body", "start", "end", "isAllDay",
         "location", "organizer", "attendees", "recurrence", "showAs",
         "sensitivity", "isCancelled", "isOrganizer", "onlineMeeting",
         "webLink", "responseStatus", "seriesMasterId", "type"];

    private static readonly JsonSerializerOptions WriteJson = new(JsonSerializerDefaults.Web)
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private readonly GraphAuthOptions _options = options.Value;
    private readonly CalendarPolicyOptions _policy = policy.Value;
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

        Event? ev = await GetGraphEventAsync(eventId.Trim(), calendarId, ct).ConfigureAwait(false);

        return ev is null
            ? throw GraphServiceException.EventNotFound(eventId, "calendar_read_event")
            : CalendarMapper.MapDetail(ev);
    }

    private async Task<Event?> GetGraphEventAsync(
        string eventId, string? calendarId, CancellationToken ct)
    {
        return (IsDefaultCalendar(calendarId), IsMe) switch
        {
            (true, true) => await client.Me.Calendar.Events[eventId].GetAsync(c =>
                c.QueryParameters.Select = EventSelect, ct).ConfigureAwait(false),
            (true, false) => await client.Users[_options.UserIdOrUpn].Calendar.Events[eventId].GetAsync(c =>
                c.QueryParameters.Select = EventSelect, ct).ConfigureAwait(false),
            (false, true) => await client.Me.Calendars[calendarId!].Events[eventId].GetAsync(c =>
                c.QueryParameters.Select = EventSelect, ct).ConfigureAwait(false),
            (false, false) => await client.Users[_options.UserIdOrUpn].Calendars[calendarId!].Events[eventId].GetAsync(c =>
                c.QueryParameters.Select = EventSelect, ct).ConfigureAwait(false)
        };
    }

    public async Task<EventDetail> CreateEventAsync(
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
        RequireDelegatedWrite();
        RequireSubject(subject);
        DateTimeOffset startValue = RequireDateTime(start, "start");
        DateTimeOffset endValue = RequireDateTime(end, "end");
        RequireValidWindow(startValue, endValue);
        ValidateAllDay(startValue, endValue, isAllDay);
        ValidateReminder(reminderMinutesBeforeStart);
        if (attendees is not null)
        {
            RecipientGuard.ValidateRecipients(attendees, ToRecipientPolicy(), "attendee");
        }

        var payload = BuildEventPayload(
            subject.Trim(), startValue, endValue, body, location, attendees,
            isAllDay, reminderMinutesBeforeStart);
        Event? created = await SendEventWriteAsync(
            Microsoft.Kiota.Abstractions.Method.POST,
            EventPath(calendarId), payload, ct).ConfigureAwait(false);

        return created is null
            ? throw GraphServiceException.GraphError(0, null, "Event creation returned no result.")
            : CalendarMapper.MapDetail(created);
    }

    public async Task<EventDetail> UpdateEventAsync(
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
        RequireDelegatedWrite();
        RequireId(eventId, "eventId");
        if (subject is null && start is null && end is null && body is null
            && location is null && attendees is null && isAllDay is null
            && reminderMinutesBeforeStart is null)
        {
            throw GraphServiceException.InvalidRequest(
                "At least one event field must be supplied.",
                "pass subject, start, end, body, location, attendees, isAllDay or reminderMinutesBeforeStart");
        }

        if (subject is not null)
        {
            RequireSubject(subject);
        }

        DateTimeOffset? startValue = start is null ? null : RequireDateTime(start, "start");
        DateTimeOffset? endValue = end is null ? null : RequireDateTime(end, "end");
        if (startValue is not null && endValue is not null)
        {
            RequireValidWindow(startValue.Value, endValue.Value);
        }

        ValidateReminder(reminderMinutesBeforeStart);
        if (attendees is not null)
        {
            RecipientGuard.ValidateRecipients(attendees, ToRecipientPolicy(), "attendee");
        }

        bool needsExistingEvent = (_policy.RequireInternalAttendees && attendees is null)
            || startValue is not null || endValue is not null || isAllDay is not null;
        Event? existing = needsExistingEvent
            ? await GetGraphEventAsync(eventId.Trim(), calendarId, ct).ConfigureAwait(false)
            : null;
        if (needsExistingEvent && existing is null)
        {
            throw GraphServiceException.EventNotFound(eventId, "calendar_update_event");
        }

        if (_policy.RequireInternalAttendees && attendees is null)
        {
            string[] existingAttendees = [..
                (existing!.Attendees ?? []).Select(a => a.EmailAddress?.Address ?? string.Empty)];
            RecipientGuard.ValidateRecipients(existingAttendees, ToRecipientPolicy(), "attendee");
        }

        bool effectiveIsAllDay = isAllDay ?? existing?.IsAllDay ?? false;
        if (startValue is not null || endValue is not null || isAllDay is not null)
        {
            DateTimeOffset effectiveStart = startValue
                ?? RequireDateTime(existing!.Start?.DateTime ?? string.Empty, "existing start");
            DateTimeOffset effectiveEnd = endValue
                ?? RequireDateTime(existing!.End?.DateTime ?? string.Empty, "existing end");
            RequireValidWindow(effectiveStart, effectiveEnd);
            ValidateAllDay(effectiveStart, effectiveEnd, effectiveIsAllDay);
        }

        var payload = BuildUpdatePayload(
            subject, startValue, endValue, body, location, attendees,
            isAllDay, effectiveIsAllDay, reminderMinutesBeforeStart);
        Event? updated = await SendEventWriteAsync(
            Microsoft.Kiota.Abstractions.Method.PATCH,
            EventPath(calendarId, eventId.Trim()), payload, ct).ConfigureAwait(false);
        return updated is null
            ? throw GraphServiceException.EventNotFound(eventId, "calendar_update_event")
            : CalendarMapper.MapDetail(updated);
    }

    private static bool IsDefaultCalendar(string? calendarId) =>
        string.IsNullOrWhiteSpace(calendarId)
        || string.Equals(calendarId.Trim(), "default", StringComparison.OrdinalIgnoreCase);

    private RecipientPolicyOptions ToRecipientPolicy() => new()
    {
        RequireInternalRecipients = _policy.RequireInternalAttendees,
        AllowedRecipientDomains = _policy.AllowedAttendeeDomains,
        AllowedRecipientAddresses = _policy.AllowedAttendeeAddresses
    };

    private async Task<Event?> SendEventWriteAsync(
        Method method, string path, IReadOnlyDictionary<string, object?> payload, CancellationToken ct)
    {
        var request = new RequestInformation
        {
            HttpMethod = method,
            UrlTemplate = "{+baseurl}" + path,
            PathParameters = new Dictionary<string, object>
            {
                ["baseurl"] = client.RequestAdapter.BaseUrl!
            }
        };
        request.Headers.Add("Accept", "application/json");
        using var stream = new MemoryStream(
            Encoding.UTF8.GetBytes(JsonSerializer.Serialize(payload, WriteJson)));
        request.SetStreamContent(stream, "application/json");
        return await client.RequestAdapter.SendAsync(
            request, Event.CreateFromDiscriminatorValue, cancellationToken: ct).ConfigureAwait(false);
    }

    private string EventPath(string? calendarId, string? eventId = null)
    {
        string ownerPath = IsMe
            ? "/me"
            : $"/users/{Uri.EscapeDataString(_options.UserIdOrUpn)}";
        string calendarPath = IsDefaultCalendar(calendarId)
            ? "/calendar/events"
            : $"/calendars/{Uri.EscapeDataString(calendarId!.Trim())}/events";
        return ownerPath + calendarPath + (eventId is null ? string.Empty : $"/{Uri.EscapeDataString(eventId)}");
    }

    private void RequireDelegatedWrite()
    {
        if (_options.AuthMode != AuthMode.Delegated)
        {
            throw GraphServiceException.AuthMisconfigured(
                "Calendar create/update requires delegated auth. Set Graph:AuthMode to Delegated.");
        }
    }

    private static Dictionary<string, object?> BuildEventPayload(
        string subject,
        DateTimeOffset start,
        DateTimeOffset end,
        string? body,
        string? location,
        IReadOnlyList<string>? attendees,
        bool isAllDay,
        int? reminderMinutesBeforeStart)
    {
        var payload = new Dictionary<string, object?>
        {
            ["subject"] = subject,
            ["start"] = ToGraphDateTime(start, isAllDay),
            ["end"] = ToGraphDateTime(end, isAllDay),
            ["isAllDay"] = isAllDay
        };

        AddOptionalPayload(payload, body, location, attendees, reminderMinutesBeforeStart);
        return payload;
    }

    private static Dictionary<string, object?> BuildUpdatePayload(
        string? subject,
        DateTimeOffset? start,
        DateTimeOffset? end,
        string? body,
        string? location,
        IReadOnlyList<string>? attendees,
        bool? isAllDay,
        bool effectiveIsAllDay,
        int? reminderMinutesBeforeStart)
    {
        var payload = new Dictionary<string, object?>();
        if (subject is not null)
        {
            payload["subject"] = subject.Trim();
        }

        if (start is not null)
        {
            payload["start"] = ToGraphDateTime(start.Value, effectiveIsAllDay);
        }

        if (end is not null)
        {
            payload["end"] = ToGraphDateTime(end.Value, effectiveIsAllDay);
        }

        if (isAllDay is not null)
        {
            payload["isAllDay"] = isAllDay.Value;
        }

        AddOptionalPayload(payload, body, location, attendees, reminderMinutesBeforeStart);
        return payload;
    }

    private static void AddOptionalPayload(
        IDictionary<string, object?> payload,
        string? body,
        string? location,
        IReadOnlyList<string>? attendees,
        int? reminderMinutesBeforeStart)
    {
        if (body is not null)
        {
            payload["body"] = new Dictionary<string, string>
            {
                ["contentType"] = "text",
                ["content"] = body
            };
        }

        if (location is not null)
        {
            payload["location"] = new Dictionary<string, string>
            {
                ["displayName"] = location
            };
        }

        if (attendees is not null)
        {
            payload["attendees"] = BuildAttendeePayload(attendees);
        }

        if (reminderMinutesBeforeStart is not null)
        {
            payload["reminderMinutesBeforeStart"] = reminderMinutesBeforeStart.Value;
            payload["isReminderOn"] = true;
        }
    }

    private static List<Dictionary<string, object>> BuildAttendeePayload(IReadOnlyList<string> attendees) =>
        [.. attendees.Select(address =>
        {
            if (string.IsNullOrWhiteSpace(address))
            {
                throw GraphServiceException.InvalidRequest(
                    "Attendee addresses must not be empty.",
                    "pass SMTP addresses such as person@example.com");
            }

            string trimmed = address.Trim();
            return new Dictionary<string, object>
            {
                ["emailAddress"] = new Dictionary<string, string>
                {
                    ["address"] = trimmed,
                    ["name"] = trimmed
                },
                ["type"] = "required"
            };
        })];

    private static Dictionary<string, string> ToGraphDateTime(DateTimeOffset value, bool isAllDay) =>
        new()
        {
            ["dateTime"] = isAllDay
                ? value.Date.ToString("yyyy-MM-dd'T'00:00:00", CultureInfo.InvariantCulture)
                : value.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture),
            ["timeZone"] = "UTC"
        };

    private static void ValidateAllDay(DateTimeOffset start, DateTimeOffset end, bool isAllDay)
    {
        if (isAllDay && (start.TimeOfDay != TimeSpan.Zero || end.TimeOfDay != TimeSpan.Zero))
        {
            throw GraphServiceException.InvalidRequest(
                "All-day event start and end must be at midnight in the supplied time zone.",
                "pass ISO values such as \"2026-09-14T00:00:00+02:00\"");
        }
    }

    private static void RequireSubject(string subject)
    {
        if (string.IsNullOrWhiteSpace(subject))
        {
            throw GraphServiceException.InvalidRequest(
                "Subject must not be empty.", "pass a title for the event");
        }
    }

    private static void RequireId(string value, string what)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw GraphServiceException.InvalidRequest(
                $"{what} must not be empty.", $"use an id from calendar_list_events or calendar_search_events");
        }
    }

    private static DateTimeOffset RequireDateTime(string value, string what)
    {
        if (string.IsNullOrWhiteSpace(value)
            || !DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeUniversal, out var parsed))
        {
            throw GraphServiceException.InvalidRequest(
                $"{what} '{value}' is not a valid date/time.",
                "use ISO format, e.g. \"2026-09-14T14:00:00Z\"");
        }

        return parsed;
    }

    private static void RequireValidWindow(DateTimeOffset start, DateTimeOffset end)
    {
        if (end <= start)
        {
            throw GraphServiceException.InvalidRequest(
                "Event end must be after event start.",
                "pass an end time later than the start time");
        }
    }

    private static void ValidateReminder(int? reminderMinutesBeforeStart)
    {
        if (reminderMinutesBeforeStart < 0)
        {
            throw GraphServiceException.InvalidRequest(
                "reminderMinutesBeforeStart must not be negative.",
                "pass zero or a positive number of minutes");
        }
    }

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
