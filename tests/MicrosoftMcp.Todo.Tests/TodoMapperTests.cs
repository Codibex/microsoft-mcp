using AwesomeAssertions;
using Microsoft.Graph.Models;

namespace MicrosoftMcp.Todo.Tests;

public sealed class TodoMapperTests
{
    [Fact]
    public void MapList_keeps_shared_flags()
    {
        var list = new TodoTaskList
        {
            Id = "l-1",
            DisplayName = "Family",
            IsOwner = false,
            IsShared = true,
            WellknownListName = WellknownListName.None
        };

        TodoListInfo mapped = TodoMapper.MapList(list);

        mapped.Id.Should().Be("l-1");
        mapped.IsShared.Should().BeTrue();
        mapped.WellknownListName.Should().Be("None");
    }

    [Fact]
    public void MapDetail_formats_due_with_timezone()
    {
        var task = new TodoTask
        {
            Id = "t-1",
            Title = "Offer",
            Status = Microsoft.Graph.Models.TaskStatus.NotStarted,
            Importance = Importance.High,
            DueDateTime = new DateTimeTimeZone { DateTime = "2026-10-15T17:00:00", TimeZone = "UTC" },
            Body = new ItemBody { Content = " notes ", ContentType = BodyType.Text },
            Categories = ["red"]
        };

        TodoTaskDetail mapped = TodoMapper.MapDetail(task, "l-1");

        mapped.DueDateTime.Should().Be("2026-10-15T17:00:00 (UTC)");
        mapped.Body.Should().Be("notes");
        mapped.Categories.Should().ContainSingle(c => c == "red");
    }

    [Fact]
    public void MapSummary_handles_missing_optionals()
    {
        var task = new TodoTask { Id = "t-2", Title = "Plain" };

        TodoTaskSummary mapped = TodoMapper.MapSummary(task, "l-1");

        mapped.DueDateTime.Should().BeNull();
        mapped.Status.Should().BeNull();
    }
}
