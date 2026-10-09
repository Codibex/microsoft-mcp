using AwesomeAssertions;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Serialization;

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

    [Fact]
    public void MapDetail_reads_untyped_checklist_from_wire()
    {
        // Real Graph responses materialize open-type checklist entries as
        // UntypedObject (no discriminator), not PlannerChecklistItem.
        var task = new PlannerTask { Id = "t-3", Title = "Wired" };
        var details = new PlannerTaskDetails
        {
            Checklist = new PlannerChecklistItems
            {
                AdditionalData = new Dictionary<string, object>
                {
                    ["chk1"] = new UntypedObject(new Dictionary<string, UntypedNode>
                    {
                        ["title"] = new UntypedString("Entwurf"),
                        ["isChecked"] = new UntypedBoolean(false),
                        ["orderHint"] = new UntypedString("a")
                    })
                }
            }
        };

        PlannerTaskDetail mapped = PlannerMapper.MapDetail(task, details, null);

        mapped.Checklist.Should().ContainSingle(c => c.Title == "Entwurf" && c.IsChecked == false);
    }
}
