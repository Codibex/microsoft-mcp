using System.Net;
using System.Text;
using System.Text.Json;
using AwesomeAssertions;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Kiota.Http.HttpClientLibrary;
using MicrosoftMcp.Common;
using NSubstitute;

namespace MicrosoftMcp.Outlook.Tests;

public sealed class GraphMailServiceTests
{
    [Fact]
    public async Task ListFolders_traverses_nested_child_folders_and_builds_paths()
    {
        var setup = Create(
            Response("""
                {"value":[{"id":"school","displayName":"School","parentFolderId":"root","totalItemCount":2,"unreadItemCount":1}]}
                """),
            Response("""
                {"value":[{"id":"kids","displayName":"Kids","parentFolderId":"school","totalItemCount":1,"unreadItemCount":1}]}
                """),
            Response("""
                {"value":[{"id":"sophie","displayName":"Sophie","parentFolderId":"kids","totalItemCount":1,"unreadItemCount":0}]}
                """),
            Response("""{"value":[]}"""));

        var folders = await setup.Service.ListFoldersAsync();

        folders.Select(folder => folder.Path).Should().Equal("School", "School/Kids", "School/Kids/Sophie");
        folders.Single(folder => folder.Id == "sophie").ParentId.Should().Be("kids");
        setup.Handler.Requests.Select(request => request.RequestUri!.AbsolutePath)
            .Should().Equal(
                "/v1.0/me/mailFolders",
                "/v1.0/me/mailFolders/school/childFolders",
                "/v1.0/me/mailFolders/kids/childFolders",
                "/v1.0/me/mailFolders/sophie/childFolders");
    }

    [Fact]
    public async Task ListFolders_follows_root_and_child_paging_links()
    {
        var setup = Create(
            Response("""
                {"value":[{"id":"school","displayName":"School"}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/me/mailFolders?$skiptoken=root-next"}
                """),
            Response("""{"value":[{"id":"archive","displayName":"Archive"}]}"""),
            Response("""
                {"value":[{"id":"kids","displayName":"Kids"}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/me/mailFolders/school/childFolders?$skiptoken=child-next"}
                """),
            Response("""{"value":[{"id":"clubs","displayName":"Clubs"}]}"""),
            Response("""{"value":[]}"""),
            Response("""{"value":[]}"""),
            Response("""{"value":[]}"""));

        var folders = await setup.Service.ListFoldersAsync();

        folders.Select(folder => folder.Path).Should().Equal(
            "School", "Archive", "School/Kids", "School/Clubs");
        setup.Handler.Requests.Should().HaveCount(7);
        setup.Handler.Requests[1].RequestUri!.Query.Should().Contain("$skiptoken=root-next");
        setup.Handler.Requests[3].RequestUri!.Query.Should().Contain("$skiptoken=child-next");
        Uri.UnescapeDataString(setup.Handler.Requests[2].RequestUri!.Query)
            .Should().Contain("$top=100");
    }

    [Fact]
    public async Task CreateFolder_with_parent_uses_child_folders_endpoint()
    {
        var setup = Create(
            Response("""{"value":[{"id":"school","displayName":"School"}]}"""),
            Response("""{"value":[{"id":"kids","displayName":"Kids","parentFolderId":"school"}]}"""),
            Response("""{"value":[]}"""),
            Response("""
                {"id":"sophie","displayName":"Sophie","parentFolderId":"kids"}
                """));

        var folder = await setup.Service.CreateFolderAsync("Sophie", "kids");

        folder.Id.Should().Be("sophie");
        folder.Path.Should().Be("School/Kids/Sophie");
        setup.Handler.Requests.Should().HaveCount(4);
        setup.Handler.Requests[^1].Method.Should().Be(HttpMethod.Post);
        setup.Handler.Requests[^1].RequestUri!.AbsolutePath
            .Should().Be("/v1.0/me/mailFolders/kids/childFolders");
    }

    [Fact]
    public async Task Move_resolves_nested_path_to_graph_folder_id()
    {
        var setup = Create(
            Response("""{"value":[{"id":"school","displayName":"School"}]}"""),
            Response("""{"value":[{"id":"kids","displayName":"Kids","parentFolderId":"school"}]}"""),
            Response("""{"value":[{"id":"sophie","displayName":"Sophie","parentFolderId":"kids"}]}"""),
            Response("""{"value":[]}"""),
            Response("""{"id":"message-1","subject":"Moved"}"""));

        var moved = await setup.Service.MoveAsync("message-1", "School/Kids/Sophie");

        moved.Id.Should().Be("message-1");
        setup.Handler.Requests[^1].RequestUri!.AbsolutePath
            .Should().Be("/v1.0/me/messages/message-1/move");
        using var payload = JsonDocument.Parse(setup.Handler.RequestBodies[^1]);
        payload.RootElement.EnumerateObject()
            .Single(property => string.Equals(property.Name, "destinationId", StringComparison.OrdinalIgnoreCase))
            .Value.GetString().Should().Be("sophie");
    }

    private static (GraphMailService Service, SequenceHandler Handler, HttpClient Client) Create(
        params HttpResponseMessage[] responses)
    {
        var handler = new SequenceHandler(responses);
        var client = new HttpClient(handler);
        var adapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: client);
        var graphClient = new GraphServiceClient(adapter);
        var service = new GraphMailService(
            graphClient,
            Options.Create(new GraphAuthOptions()),
            Options.Create(new OutlookPolicyOptions()));
        return (service, handler, client);
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