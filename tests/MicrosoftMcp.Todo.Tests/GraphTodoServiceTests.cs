using System.Net;
using System.Text;
using AwesomeAssertions;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Kiota.Http.HttpClientLibrary;
using MicrosoftMcp.Common;
using NSubstitute;

namespace MicrosoftMcp.Todo.Tests;

/// <summary>Guards the documented To Do paths and the known backend quirk:
/// the To Do service rejects $select with 400 RequestBroker--ParseUri, so
/// no read sends it ($top + $filter on status are supported).</summary>
public sealed class GraphTodoServiceTests
{
    private static GraphTodoService Create(SequenceHandler handler)
    {
        var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        return new GraphTodoService(
            new GraphServiceClient(requestAdapter),
            Options.Create(new GraphAuthOptions()));
    }

    [Fact]
    public async Task ListLists_uses_documented_path_without_select()
    {
        var handler = new SequenceHandler(
            Response("""{"value":[{"id":"l-1","displayName":"Tasks","isOwner":true,"isShared":false,"wellknownListName":"defaultList"}]}"""));
        var service = Create(handler);

        var lists = await service.ListListsAsync();

        lists.Should().ContainSingle(l => l.Id == "l-1");
        handler.Requests.Should().HaveCount(1);
        handler.Requests[0].RequestUri!.AbsolutePath.Should().Be("/v1.0/me/todo/lists");
        handler.Requests[0].RequestUri!.Query.Should().NotContain("$select");
    }

    [Fact]
    public async Task ListTasks_filters_completed_without_select_or_orderby()
    {
        var handler = new SequenceHandler(
            Response("""{"id":"l-1","displayName":"Tasks","wellknownListName":"none"}"""),
            Response("""
                {"value":[
                    {"id":"t-old","title":"Old","status":"notStarted","createdDateTime":"2026-09-01T10:00:00Z"},
                    {"id":"t-new","title":"New","status":"notStarted","createdDateTime":"2026-10-01T10:00:00Z"},
                    {"id":"t-done","title":"Done","status":"completed","createdDateTime":"2026-10-02T10:00:00Z"}
                ]}
                """));
        var service = Create(handler);

        var tasks = await service.ListTasksAsync("l-1");

        tasks.Select(t => t.Id).Should().Equal("t-new", "t-old");
        handler.Requests.Should().HaveCount(2);
        handler.Requests[0].RequestUri!.AbsolutePath.Should().Be("/v1.0/me/todo/lists/l-1");
        handler.Requests[0].RequestUri!.Query.Should().NotContain("$select");
        var listCall = handler.Requests.Last();
        listCall.RequestUri!.AbsolutePath.Should().Be("/v1.0/me/todo/lists/l-1/tasks");
        string query = Uri.UnescapeDataString(listCall.RequestUri!.Query);
        query.Should().Contain("$top=");
        query.Should().Contain("$filter=status ne 'completed'");
        query.Should().NotContain("$select");
        query.Should().NotContain("$orderby");
    }

    [Fact]
    public async Task ListTasks_without_listId_resolves_default_list()
    {
        var handler = new SequenceHandler(
            Response("""{"value":[{"id":"l-1","displayName":"Tasks","wellknownListName":"defaultList"},{"id":"l-2","displayName":"Family","wellknownListName":"none"}]}"""),
            Response("""{"value":[{"id":"t-1","title":"Hi","status":"notStarted"}]}"""));
        var service = Create(handler);

        var tasks = await service.ListTasksAsync(includeCompleted: true);

        tasks.Should().ContainSingle(t => t.Id == "t-1");
        handler.Requests.Should().HaveCount(2);
        handler.Requests[0].RequestUri!.AbsolutePath.Should().Be("/v1.0/me/todo/lists");
        handler.Requests[1].RequestUri!.AbsolutePath.Should().Be("/v1.0/me/todo/lists/l-1/tasks");
        Uri.UnescapeDataString(handler.Requests[1].RequestUri!.Query).Should().NotContain("$filter");
    }

    [Fact]
    public async Task ListTasks_pages_before_sorting_so_newer_tasks_win()
    {
        var handler = new SequenceHandler(
            Response("""{"id":"l-1","displayName":"Tasks","wellknownListName":"none"}"""),
            Response("""
                {"value":[
                    {"id":"t-old","title":"Old","status":"notStarted","createdDateTime":"2026-09-01T10:00:00Z"}
                ],"@odata.nextLink":"https://graph.microsoft.com/v1.0/me/todo/lists/l-1/tasks?$skiptoken=next"}
                """),
            Response("""
                {"value":[
                    {"id":"t-new","title":"New","status":"notStarted","createdDateTime":"2026-10-01T10:00:00Z"}
                ]}
                """));
        var service = Create(handler);

        // take=1: stopping at the first page would return t-old.
        var tasks = await service.ListTasksAsync("l-1", top: 1);

        tasks.Select(t => t.Id).Should().Equal("t-new");
        handler.Requests.Should().HaveCount(3);
        handler.Requests[2].RequestUri!.Query.Should().Contain("$skiptoken=next");
    }

    [Fact]
    public async Task AddTask_posts_documented_wire_shape()
    {
        var handler = new SequenceHandler(
            Response("""{"id":"l-1","displayName":"Tasks","wellknownListName":"none"}"""),
            Response("""{"id":"t-9","title":"Offer","status":"notStarted"}"""));
        var service = Create(handler);

        var added = await service.AddTaskAsync(
            "Offer", "l-1", dueDateTime: "2026-10-15T17:00:00Z", importance: "high", body: "notes");

        added.Id.Should().Be("t-9");
        var post = handler.Requests.Last();
        post.Method.Should().Be(HttpMethod.Post);
        post.RequestUri!.AbsolutePath.Should().Be("/v1.0/me/todo/lists/l-1/tasks");
        string body = handler.RequestBodies.Last();
        body.Should().Contain("\"title\":\"Offer\"");
        body.Should().Contain("\"importance\":\"high\"");
        body.Should().NotContain("completed");
        body.Should().Contain("2026-10-15");
        body.Should().Contain("\"timeZone\":\"UTC\"");
    }

    [Fact]
    public async Task CompleteTask_patches_status_completed()
    {
        var handler = new SequenceHandler(
            Response("""{"id":"l-1","displayName":"Tasks","wellknownListName":"none"}"""),
            Response("""{"id":"t-1","title":"Hi","status":"completed"}"""));
        var service = Create(handler);

        var done = await service.CompleteTaskAsync("t-1", "l-1");

        done.Status.Should().Be("Completed");
        handler.Requests.Should().HaveCount(2);
        handler.Requests[1].Method.Should().Be(new HttpMethod("PATCH"));
        handler.Requests[1].RequestUri!.AbsolutePath.Should().Be("/v1.0/me/todo/lists/l-1/tasks/t-1");
        handler.RequestBodies[1].Should().Contain("\"status\":\"completed\"");
    }

    [Fact]
    public async Task CompleteTask_without_listId_discovers_owner_and_patches()
    {
        var handler = new SequenceHandler(
            Response("""{"value":[{"id":"l-1","displayName":"Tasks","wellknownListName":"defaultList"}]}"""),
            Response("""{"id":"t-1","title":"Hi","status":"notStarted"}"""),
            Response("""{"id":"t-1","title":"Hi","status":"completed"}"""));
        var service = Create(handler);

        var done = await service.CompleteTaskAsync("t-1");

        done.Status.Should().Be("Completed");
        handler.Requests.Should().HaveCount(3);
        handler.Requests[1].RequestUri!.AbsolutePath.Should().Be("/v1.0/me/todo/lists/l-1/tasks/t-1");
        handler.Requests[1].RequestUri!.Query.Should().NotContain("$select");
        handler.Requests[2].Method.Should().Be(new HttpMethod("PATCH"));
    }

    [Fact]
    public async Task CompleteTask_unknown_id_maps_to_todo_task_not_found()
    {
        var handler = new SequenceHandler(
            Response("""{"id":"l-1","displayName":"Tasks","wellknownListName":"none"}"""),
            new HttpResponseMessage(HttpStatusCode.NotFound)
            {
                Content = new StringContent(
                    """{"error":{"code":"ErrorItemNotFound","message":"Not found"}}""",
                    Encoding.UTF8,
                    "application/json")
            });
        var service = Create(handler);

        var ex = await Assert.ThrowsAsync<GraphServiceException>(
            () => service.CompleteTaskAsync("nope", "l-1"));

        ex.Code.Should().Be("todo-task-not-found");
    }

    [Fact]
    public async Task CompleteTask_invalid_list_maps_to_todo_list_not_found()
    {
        var handler = new SequenceHandler(
            new HttpResponseMessage(HttpStatusCode.NotFound)
            {
                Content = new StringContent(
                    """{"error":{"code":"ErrorItemNotFound","message":"Not found"}}""",
                    Encoding.UTF8,
                    "application/json")
            });
        var service = Create(handler);

        // The list is validated before patching, so an invalid list reports
        // [todo-list-not-found] even though the PATCH path would also 404.
        var ex = await Assert.ThrowsAsync<GraphServiceException>(
            () => service.CompleteTaskAsync("t-1", "no-list"));

        ex.Code.Should().Be("todo-list-not-found");
        handler.Requests.Should().HaveCount(1);
        handler.Requests[0].Method.Should().Be(HttpMethod.Get);
        handler.Requests[0].RequestUri!.AbsolutePath.Should().Be("/v1.0/me/todo/lists/no-list");
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
