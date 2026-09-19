using System.Text.Json;
using AwesomeAssertions;
using Microsoft.Extensions.Logging.Abstractions;
using ModelContextProtocol.Protocol;
using NSubstitute;

namespace MicrosoftMcp.Calendar.Tests;

internal static class ToolResults
{
    private static readonly JsonSerializerOptions Read =
        new(ToolResult.Json) { PropertyNameCaseInsensitive = true };

    internal static T Ok<T>(CallToolResult result)
    {
        Assert.False(result.IsError is true);
        return JsonSerializer.Deserialize<T>(ToolResult.ReadText(result), Read)
            ?? throw new InvalidOperationException("Tool result was null JSON.");
    }

    internal static string Fail(CallToolResult result)
    {
        Assert.True(result.IsError is true);
        return ToolResult.ReadText(result);
    }
}

public sealed class CalendarToolsTests
{
    private static CalendarTools Create(IGraphCalendarService calendar) =>
        new(calendar, NullLogger<CalendarTools>.Instance);

    [Fact]
    public async Task Browse_calendars_and_window_with_fake()
    {
        var tools = Create(new FakeGraphCalendarService());

        var calendars = ToolResults.Ok<List<CalendarInfo>>(await tools.calendar_list_calendars());
        calendars.Should().HaveCount(2);
        calendars.Should().Contain(c => c.IsDefault);

        var upcoming = ToolResults.Ok<List<EventSummary>>(await tools.calendar_list_events());
        upcoming.Should().Contain(e => e.Id == "e-standup");
        upcoming.Should().NotContain(e => e.Id == "e-old");

        var birthdays = ToolResults.Ok<List<EventSummary>>(
            await tools.calendar_list_events("cal-birth"));
        birthdays.Should().ContainSingle(e => e.Id == "e-bday");
    }

    [Fact]
    public async Task Search_and_read_with_fake()
    {
        var tools = Create(new FakeGraphCalendarService());

        var found = ToolResults.Ok<List<EventSummary>>(await tools.calendar_search_events("standup"));
        found.Should().ContainSingle(e => e.Id == "e-standup");

        var detail = ToolResults.Ok<EventDetail>(await tools.calendar_read_event("e-standup"));
        detail.Subject.Should().Be("Daily Standup");
        detail.Attendees.Should().HaveCount(1);
    }

    [Fact]
    public async Task Create_and_update_with_fake()
    {
        var fake = new FakeGraphCalendarService();
        var tools = Create(fake);

        var created = ToolResults.Ok<EventDetail>(await tools.calendar_create_event(
            "Planning",
            "2026-09-14T14:00:00Z",
            "2026-09-14T15:00:00Z",
            body: "Agenda",
            location: "Room 3",
            attendees: ["guest@example.com"],
            reminderMinutesBeforeStart: 15));
        created.Subject.Should().Be("Planning");
        created.Body.Should().Be("Agenda");
        created.Location.Should().Be("Room 3");
        created.Attendees.Should().ContainSingle(a => a.Address == "guest@example.com");
        fake.GetReminderMinutesBeforeStart(created.Id).Should().Be(15);

        var updated = ToolResults.Ok<EventDetail>(await tools.calendar_update_event(
            created.Id,
            subject: "Moved",
            body: "Updated agenda",
            location: "Room 4",
            attendees: ["organizer@example.com"],
            reminderMinutesBeforeStart: 30));
        updated.Subject.Should().Be("Moved");
        updated.Location.Should().Be("Room 4");
        updated.Body.Should().Be("Updated agenda");
        updated.Attendees.Should().ContainSingle(a => a.Address == "organizer@example.com");
        fake.GetReminderMinutesBeforeStart(created.Id).Should().Be(30);
    }

    [Fact]
    public async Task Tool_errors_carry_codes_and_hints()
    {
        var tools = Create(new FakeGraphCalendarService());

        ToolResults.Fail(await tools.calendar_read_event("nope")).Should().Contain("[event-not-found]");
        ToolResults.Fail(await tools.calendar_read_event("  ")).Should().Contain("[invalid-request]");
        ToolResults.Fail(await tools.calendar_search_events("  ")).Should().Contain("[invalid-request]");
        ToolResults.Fail(await tools.calendar_list_events("no-cal")).Should().Contain("[calendar-not-found]");
        ToolResults.Fail(await tools.calendar_list_events(timeMin: "kein-datum"))
            .Should().Contain("[invalid-request]");
    }

    [Fact]
    public async Task Unexpected_backend_failures_are_mapped_not_leaked()
    {
        IGraphCalendarService failing = Substitute.For<IGraphCalendarService>();
        failing.ListCalendarsAsync(Arg.Any<CancellationToken>())
            .Returns<Task<IReadOnlyList<CalendarInfo>>>(_ => throw new HttpRequestException("no route"));
        var tools = Create(failing);

        var text = ToolResults.Fail(await tools.calendar_list_calendars());
        text.Should().Contain("[service-unavailable]");
        text.Should().Contain("Next:");
    }

    [Fact]
    public async Task Delegation_wiring_with_substitute()
    {
        IGraphCalendarService calendar = Substitute.For<IGraphCalendarService>();
        calendar.ListCalendarsAsync(Arg.Any<CancellationToken>()).Returns(
            [new CalendarInfo("c1", "Kalender", null, true, true)]);
        var tools = Create(calendar);

        ToolResults.Ok<List<CalendarInfo>>(await tools.calendar_list_calendars())
            .Should().ContainSingle(c => c.Id == "c1");
    }
}
