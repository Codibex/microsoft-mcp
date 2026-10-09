using System.Net;
using System.Text;
using AwesomeAssertions;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Kiota.Http.HttpClientLibrary;
using MicrosoftMcp.Common;
using NSubstitute;

namespace MicrosoftMcp.Planner.Tests;

/// <summary>Guards the documented Planner paths (planner-concept-overview,
/// planneruser-list-tasks, plannergroup-list-plans, plannerplan-list-tasks)
/// and the open-type assignments dictionary (keys are user ids).</summary>
public sealed class GraphPlannerServiceTests
{
    private static GraphPlannerService Create(SequenceHandler handler)
    {
        var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        return new GraphPlannerService(
            new GraphServiceClient(requestAdapter),
            Options.Create(new GraphAuthOptions()));
    }

    [Fact]
    public async Task ListMyTasks_uses_documented_path()
    {
        var handler = new SequenceHandler(
            Response("""{"value":[{"id":"task-1","planId":"plan-1","bucketId":"b-1","title":"Offer","percentComplete":0}]}"""));
        var service = Create(handler);

        var tasks = await service.ListMyTasksAsync();

        tasks.Should().ContainSingle(t => t.Id == "task-1");
        handler.Requests.Should().HaveCount(1);
        handler.Requests[0].RequestUri!.AbsolutePath.Should().Be("/v1.0/me/planner/tasks");
        Uri.UnescapeDataString(handler.Requests[0].RequestUri!.Query).Should().Contain("$top");
    }

    [Fact]
    public async Task ListPlans_uses_group_path()
    {
        var handler = new SequenceHandler(
            Response("""{"value":[{"id":"plan-1","title":"Launch","owner":"g-eng"}]}"""));
        var service = Create(handler);

        var plans = await service.ListPlansAsync("g-eng");

        plans.Should().ContainSingle(p => p.Id == "plan-1");
        handler.Requests.Should().HaveCount(1);
        handler.Requests[0].RequestUri!.AbsolutePath.Should().Be("/v1.0/groups/g-eng/planner/plans");
    }

    [Fact]
    public async Task ListPlans_follows_paging_until_take()
    {
        var handler = new SequenceHandler(
            Response("""{"value":[{"id":"plan-1","title":"A"}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/groups/g-eng/planner/plans?$skiptoken=next"}"""),
            Response("""{"value":[{"id":"plan-2","title":"B"}]}"""));
        var service = Create(handler);

        var plans = await service.ListPlansAsync("g-eng");

        plans.Select(p => p.Id).Should().Equal("plan-1", "plan-2");
        handler.Requests.Should().HaveCount(2);
        handler.Requests[1].RequestUri!.Query.Should().Contain("$skiptoken=next");
    }

    [Fact]
    public async Task ListPlans_unknown_group_maps_to_group_not_found()
    {
        var handler = new SequenceHandler(
            new HttpResponseMessage(HttpStatusCode.NotFound)
            {
                Content = new StringContent(
                    """{"error":{"code":"UnknownError","message":"Group not found"}}""",
                    Encoding.UTF8,
                    "application/json")
            });
        var service = Create(handler);

        var ex = await Assert.ThrowsAsync<GraphServiceException>(
            () => service.ListPlansAsync("no-group"));

        ex.Code.Should().Be("group-not-found");
    }

    [Fact]
    public async Task ListPlanTasks_uses_plan_path()
    {
        var handler = new SequenceHandler(
            Response("""{"value":[{"id":"task-1","planId":"plan-1","bucketId":"b-1","title":"Offer","percentComplete":50}]}"""));
        var service = Create(handler);

        var tasks = await service.ListPlanTasksAsync("plan-1");

        tasks.Should().ContainSingle(t => t.PercentComplete == 50);
        handler.Requests.Should().HaveCount(1);
        handler.Requests[0].RequestUri!.AbsolutePath.Should().Be("/v1.0/planner/plans/plan-1/tasks");
    }

    [Fact]
    public async Task ReadTask_merges_task_details_and_bucket_name()
    {
        var handler = new SequenceHandler(
            Response("""
                {"id":"task-1","planId":"plan-1","bucketId":"b-1","title":"Offer",
                 "assignments":{"user-alice":{"@odata.type":"#microsoft.graph.plannerAssignment"}}}
                """),
            Response("""
                {"description":"  Angebot an Contoso  ",
                 "checklist":{"chk1":{"@odata.type":"#microsoft.graph.plannerChecklistItem",
                    "title":"Entwurf","isChecked":false,"orderHint":"a"}}}
                """),
            Response("""{"id":"b-1","name":"To do"}"""));
        var service = Create(handler);

        var detail = await service.ReadTaskAsync("task-1");

        detail.Description.Should().Be("Angebot an Contoso");
        detail.BucketName.Should().Be("To do");
        detail.AssigneeIds.Should().Contain("user-alice");
        detail.Checklist.Should().ContainSingle(c => c.Title == "Entwurf" && c.IsChecked == false);
        handler.Requests.Should().HaveCount(3);
        handler.Requests[0].RequestUri!.AbsolutePath.Should().Be("/v1.0/planner/tasks/task-1");
        handler.Requests[1].RequestUri!.AbsolutePath.Should().Be("/v1.0/planner/tasks/task-1/details");
        handler.Requests[2].RequestUri!.AbsolutePath.Should().Be("/v1.0/planner/buckets/b-1");
    }

    [Fact]
    public async Task ReadTask_unknown_id_maps_to_planner_task_not_found()
    {
        var handler = new SequenceHandler(
            new HttpResponseMessage(HttpStatusCode.NotFound)
            {
                Content = new StringContent(
                    """{"error":{"code":"UnknownError","message":"Task not found"}}""",
                    Encoding.UTF8,
                    "application/json")
            });
        var service = Create(handler);

        var ex = await Assert.ThrowsAsync<GraphServiceException>(
            () => service.ReadTaskAsync("nope"));

        ex.Code.Should().Be("planner-task-not-found");
    }

    private static HttpResponseMessage Response(string json) => new(HttpStatusCode.OK)
    {
        Content = new StringContent(json, Encoding.UTF8, "application/json")
    };

    private sealed class SequenceHandler(params HttpResponseMessage[] responses) : HttpMessageHandler
    {
        private readonly Queue<HttpResponseMessage> _responses = new(responses);

        public List<HttpRequestMessage> Requests { get; } = [];

        public List<string> RequestBodies { get; } = [];

        protected override async Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request,
            CancellationToken cancellationToken)
        {
            Requests.Add(request);
            RequestBodies.Add(request.Content is null
                ? string.Empty
                : await request.Content.ReadAsStringAsync(cancellationToken));
            return _responses.Dequeue();
        }
    }
}
