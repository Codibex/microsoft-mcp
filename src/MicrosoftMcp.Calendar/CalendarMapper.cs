using Microsoft.Graph.Models;

namespace MicrosoftMcp.Calendar;

/// <summary>Pure mapping helpers (Graph models to DTOs). Internal but unit-tested.</summary>
internal static class CalendarMapper
{
    internal static CalendarInfo MapCalendar(Microsoft.Graph.Models.Calendar c) => new(
        c.Id ?? string.Empty,
        c.Name ?? string.Empty,
        c.Color?.ToString(),
        c.IsDefaultCalendar ?? false,
        c.CanEdit ?? false);

    internal static EventSummary MapSummary(Event e) => new(
        e.Id ?? string.Empty,
        e.Subject ?? string.Empty,
        e.Start?.DateTime,
        e.Start?.TimeZone,
        e.End?.DateTime,
        e.End?.TimeZone,
        e.IsAllDay ?? false,
        e.Location?.DisplayName,
        e.Organizer?.EmailAddress?.Address,
        e.Attendees?.Count ?? 0,
        e.ShowAs?.ToString(),
        e.IsCancelled ?? false,
        Truncate(e.BodyPreview, 300),
        e.WebLink);

    internal static EventDetail MapDetail(Event e) => new(
        e.Id ?? string.Empty,
        e.Subject ?? string.Empty,
        e.Start?.DateTime,
        e.Start?.TimeZone,
        e.End?.DateTime,
        e.End?.TimeZone,
        e.IsAllDay ?? false,
        e.Location?.DisplayName,
        e.Organizer?.EmailAddress?.Address,
        [.. (e.Attendees ?? []).Select(a => new AttendeeDto(
            a.EmailAddress?.Name ?? string.Empty,
            a.EmailAddress?.Address ?? string.Empty,
            a.Status?.Response?.ToString()))],
        DescribeRecurrence(e),
        e.OnlineMeeting?.JoinUrl,
        e.Sensitivity?.ToString(),
        e.IsOrganizer ?? false,
        e.ShowAs?.ToString(),
        e.IsCancelled ?? false,
        Truncate(e.BodyPreview, 500),
        Truncate(e.Body?.Content, 8000),
        e.WebLink);

    internal static string? DescribeRecurrence(Event e)
    {
        var pattern = e.Recurrence?.Pattern;
        if (pattern?.Type is null)
        {
            return null;
        }

        string range = e.Recurrence?.Range switch
        {
            { Type: not null, StartDate: not null, EndDate: not null } r =>
                $" from {r.StartDate} until {r.EndDate}",
            { Type: not null, StartDate: not null } r =>
                $" from {r.StartDate}",
            _ => string.Empty
        };
        return $"{pattern.Type} every {pattern.Interval ?? 1}{range}".Trim();
    }

    internal static string? Truncate(string? s, int max) =>
        s is null || s.Length <= max ? s : s[..max] + "…[truncated]";
}
