using System.Text.Json;
using AwesomeAssertions;
using Microsoft.Extensions.Logging.Abstractions;
using ModelContextProtocol.Protocol;
using NSubstitute;

namespace MicrosoftMcp.Planner.Tests;

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

public sealed class PlannerToolsTests
{
    private static PlannerTools Create(IGraphPlannerService planner) =>
        new(planner, NullLogger<PlannerTools>.Instance);

    [Fact]
    public async Task Browse_my_tasks_plans_and_details_with_fake()
    {
        var tools = Create(new FakeGraphPlannerService());

        var mine = ToolResults.Ok<List<PlannerTaskSummary>>(await tools.planner_list_my_tasks());
        mine.Should().ContainSingle(t => t.Id == "task-1");

        var plans = ToolResults.Ok<List<PlannerPlanInfo>>(await tools.planner_list_plans("g-eng"));
        plans.Should().HaveCount(2);

        var tasks = ToolResults.Ok<List<PlannerTaskSummary>>(await tools.planner_list_plan_tasks("plan-1"));
        tasks.Should().HaveCount(2);

        var detail = ToolResults.Ok<PlannerTaskDetail>(await tools.planner_read_task("task-1"));
        detail.BucketName.Should().Be("To do");
        detail.Description.Should().Contain("Contoso");
        detail.Checklist.Should().ContainSingle(c => c.Title == "Entwurf");
        detail.AssigneeIds.Should().Contain("user-alice");
    }

    [Fact]
    public async Task Tool_errors_carry_codes_and_hints()
    {
        var tools = Create(new FakeGraphPlannerService());

        ToolResults.Fail(await tools.planner_list_plans("no-group")).Should().Contain("[group-not-found]");
        ToolResults.Fail(await tools.planner_list_plans("  ")).Should().Contain("[invalid-request]");
        ToolResults.Fail(await tools.planner_list_plan_tasks("no-plan")).Should().Contain("[planner-plan-not-found]");
        ToolResults.Fail(await tools.planner_read_task("nope")).Should().Contain("[planner-task-not-found]");
    }

    [Fact]
    public async Task Unexpected_backend_failures_are_mapped_not_leaked()
    {
        IGraphPlannerService failing = Substitute.For<IGraphPlannerService>();
        failing.ListMyTasksAsync(Arg.Any<int>(), Arg.Any<CancellationToken>())
            .Returns<Task<IReadOnlyList<PlannerTaskSummary>>>(_ => throw new HttpRequestException("no route"));
        var tools = Create(failing);

        var text = ToolResults.Fail(await tools.planner_list_my_tasks());
        text.Should().Contain("[service-unavailable]");
        text.Should().Contain("Next:");
    }

    [Fact]
    public async Task Delegation_wiring_with_substitute()
    {
        IGraphPlannerService planner = Substitute.For<IGraphPlannerService>();
        planner.ListPlansAsync("g-1", Arg.Any<int>(), Arg.Any<CancellationToken>()).Returns(
            [new PlannerPlanInfo("p1", "Launch", "g-1", DateTimeOffset.UtcNow)]);
        var tools = Create(planner);

        ToolResults.Ok<List<PlannerPlanInfo>>(await tools.planner_list_plans("g-1"))
            .Should().ContainSingle(p => p.Id == "p1");
    }
}
