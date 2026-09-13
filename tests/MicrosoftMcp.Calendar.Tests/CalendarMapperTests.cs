using AwesomeAssertions;
using Microsoft.Graph.Models;

namespace MicrosoftMcp.Calendar.Tests;

public sealed class CalendarMapperTests
{
    [Fact]
    public void MapCalendar_is_null_safe()
    {
        var info = CalendarMapper.MapCalendar(new Microsoft.Graph.Models.Calendar());

        info.Id.Should().BeEmpty();
        info.IsDefault.Should().BeFalse();
    }

    [Fact]
    public void MapSummary_maps_times_and_flags()
    {
        var summary = CalendarMapper.MapSummary(new Event
        {
            Id = "e1",
            Subject = "Sync",
            Start = new DateTimeTimeZone { DateTime = "2026-09-14T10:00:00", TimeZone = "UTC" },
            IsAllDay = false,
            Location = new Location { DisplayName = "Raum 1" },
            IsCancelled = false
        });

        summary.Start.Should().Be("2026-09-14T10:00:00");
        summary.StartTimeZone.Should().Be("UTC");
        summary.Location.Should().Be("Raum 1");
    }

    [Fact]
    public void MapDetail_maps_attendees_and_truncates_body()
    {
        var detail = CalendarMapper.MapDetail(new Event
        {
            Id = "e1",
            Attendees =
            [
                new Attendee
                {
                    EmailAddress = new EmailAddress { Name = "A", Address = "a@x.y" }
                }
            ],
            Body = new ItemBody { Content = new string('z', 9000) }
        });

        detail.Attendees.Should().HaveCount(1);
        detail.Attendees[0].Address.Should().Be("a@x.y");
        detail.Body.Should().HaveLength(8000 + "…[truncated]".Length);
    }

    [Fact]
    public void DescribeRecurrence_returns_null_without_pattern() =>
        CalendarMapper.DescribeRecurrence(new Event()).Should().BeNull();

    [Fact]
    public void DescribeRecurrence_summarizes_pattern_and_range()
    {
        var text = CalendarMapper.DescribeRecurrence(new Event
        {
            Recurrence = new PatternedRecurrence
            {
                Pattern = new RecurrencePattern
                {
                    Type = RecurrencePatternType.Weekly,
                    Interval = 2
                },
                Range = new RecurrenceRange
                {
                    Type = RecurrenceRangeType.EndDate,
                    StartDate = new DateOnly(2026, 9, 1),
                    EndDate = new DateOnly(2026, 12, 31)
                }
            }
        });

        text.Should().Contain("Weekly");
        text.Should().Contain("2026-12-31");
    }
}
