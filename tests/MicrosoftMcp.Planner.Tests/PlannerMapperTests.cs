using AwesomeAssertions;
using Microsoft.Graph.Models;

namespace MicrosoftMcp.Planner.Tests;

public sealed class PlannerMapperTests
{
    [Fact]
    public void MapSummary_orders_assignees()
    {
        var task = new PlannerTask
        {
            Id = "t-1",
            PlanId = "p-1",
            BucketId = "b-1",
            Title = "Offer",
            PercentComplete = 0,
            Assignments = new PlannerAssignments
            {
                AdditionalData = new Dictionary<string, object>
                {
                    ["user-bob"] = new PlannerAssignment(),
                    ["user-alice"] = new PlannerAssignment()
                }
            }
        };

        PlannerTaskSummary mapped = PlannerMapper.MapSummary(task);

        mapped.AssigneeIds.Should().Equal("user-alice", "user-bob");
    }

    [Fact]
    public void MapDetail_keeps_description_and_bucket()
    {
        var task = new PlannerTask { Id = "t-1", PlanId = "p-1", BucketId = "b-1", Title = "Offer" };
        var details = new PlannerTaskDetails
        {
            Description = "  Angebot an Contoso  ",
            Checklist = new PlannerChecklistItems
            {
                AdditionalData = new Dictionary<string, object>
                {
                    ["c1"] = new PlannerChecklistItem { Title = "Entwurf", IsChecked = false, OrderHint = "a" }
                }
            }
        };

        PlannerTaskDetail mapped = PlannerMapper.MapDetail(task, details, "To do");

        mapped.Description.Should().Be("Angebot an Contoso");
        mapped.BucketName.Should().Be("To do");
        mapped.Checklist.Should().ContainSingle(c => c.Title == "Entwurf");
    }

    [Fact]
    public void MapDetail_handles_missing_details()
    {
        var task = new PlannerTask { Id = "t-2", Title = "Plain" };

        PlannerTaskDetail mapped = PlannerMapper.MapDetail(task, null, null);

        mapped.Description.Should().BeNull();
        mapped.Checklist.Should().BeEmpty();
    }
}
