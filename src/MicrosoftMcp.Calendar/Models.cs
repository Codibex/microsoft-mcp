namespace MicrosoftMcp.Calendar;

public sealed record CalendarInfo(
    string Id,
    string Name,
    string? Color,
    bool IsDefault,
    bool CanEdit);

public sealed record AttendeeDto(string Name, string Address, string? Response);

public sealed record EventSummary(
    string Id,
    string Subject,
    string? Start,
    string? StartTimeZone,
    string? End,
    string? EndTimeZone,
    bool IsAllDay,
    string? Location,
    string? Organizer,
    int AttendeeCount,
    string? ShowAs,
    bool IsCancelled,
    string? Preview,
    string? WebLink);

public sealed record EventDetail(
    string Id,
    string Subject,
    string? Start,
    string? StartTimeZone,
    string? End,
    string? EndTimeZone,
    bool IsAllDay,
    string? Location,
    string? Organizer,
    IReadOnlyList<AttendeeDto> Attendees,
    string? Recurrence,
    string? OnlineMeetingUrl,
    string? Sensitivity,
    bool IsOrganizer,
    string? ShowAs,
    bool IsCancelled,
    string? BodyPreview,
    string? Body,
    string? WebLink);
