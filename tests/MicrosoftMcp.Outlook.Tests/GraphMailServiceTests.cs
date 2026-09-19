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
                {"value":[{"id":"school","displayName":"School","parentFolderId":"root","childFolderCount":1,"totalItemCount":2,"unreadItemCount":1}]}
                """),
            Response("""
                {"value":[{"id":"kids","displayName":"Kids","parentFolderId":"school","childFolderCount":1,"totalItemCount":1,"unreadItemCount":1}]}
                """),
            Response("""
                {"value":[{"id":"sophie","displayName":"Sophie","parentFolderId":"kids","childFolderCount":0,"totalItemCount":1,"unreadItemCount":0}]}
                """));

        var folders = await setup.Service.ListFoldersAsync();

        folders.Select(folder => folder.Path).Should().Equal("School", "School/Kids", "School/Kids/Sophie");
        folders.Single(folder => folder.Id == "sophie").ParentId.Should().Be("kids");
        setup.Handler.Requests.Select(request => request.RequestUri!.AbsolutePath)
            .Should().Equal(
                "/v1.0/me/mailFolders",
                "/v1.0/me/mailFolders/school/childFolders",
            "/v1.0/me/mailFolders/kids/childFolders");
    }

    [Fact]
    public async Task ListFolders_follows_root_and_child_paging_links()
    {
        var setup = Create(
            Response("""
                {"value":[{"id":"school","displayName":"School","childFolderCount":1}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/me/mailFolders?$skiptoken=root-next"}
                """),
            Response("""{"value":[{"id":"archive","displayName":"Archive","childFolderCount":0}]}"""),
            Response("""
                {"value":[{"id":"kids","displayName":"Kids","childFolderCount":1}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/me/mailFolders/school/childFolders?$skiptoken=child-next"}
                """),
            Response("""{"value":[{"id":"clubs","displayName":"Clubs","childFolderCount":0}]}"""),
            Response("""{"value":[]}"""));

        var folders = await setup.Service.ListFoldersAsync();

        folders.Select(folder => folder.Path).Should().Equal(
            "School", "Archive", "School/Kids", "School/Clubs");
        setup.Handler.Requests.Should().HaveCount(5);
        setup.Handler.Requests[1].RequestUri!.Query.Should().Contain("$skiptoken=root-next");
        setup.Handler.Requests[3].RequestUri!.Query.Should().Contain("$skiptoken=child-next");
        Uri.UnescapeDataString(setup.Handler.Requests[2].RequestUri!.Query)
            .Should().Contain("$top=100");
    }

    [Fact]
    public async Task User_folder_traversal_follows_root_and_child_paging_links()
    {
        var setup = Create(
            new GraphAuthOptions { UserIdOrUpn = "user-1" },
            Response("""
                {"value":[{"id":"school","displayName":"School","childFolderCount":1}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/users/user-1/mailFolders?$skiptoken=root-next"}
                """),
            Response("""{"value":[{"id":"archive","displayName":"Archive","childFolderCount":0}]}"""),
            Response("""
                {"value":[{"id":"kids","displayName":"Kids","childFolderCount":0}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/users/user-1/mailFolders/school/childFolders?$skiptoken=child-next"}
                """),
            Response("""{"value":[{"id":"clubs","displayName":"Clubs","childFolderCount":0}]}"""));

        var folders = await setup.Service.ListFoldersAsync();

        folders.Select(folder => folder.Path).Should().Equal("School", "Archive", "School/Kids", "School/Clubs");
        setup.Handler.Requests.Should().HaveCount(4);
        setup.Handler.Requests[0].RequestUri!.AbsolutePath.Should().Be("/v1.0/users/user-1/mailFolders");
        setup.Handler.Requests[1].RequestUri!.Query.Should().Contain("$skiptoken=root-next");
        setup.Handler.Requests[2].RequestUri!.AbsolutePath
            .Should().Be("/v1.0/users/user-1/mailFolders/school/childFolders");
        setup.Handler.Requests[3].RequestUri!.Query.Should().Contain("$skiptoken=child-next");
    }

    [Fact]
    public async Task ListAttachments_follows_next_link()
    {
        var setup = Create(
            Response("""
                {"value":[{"@odata.type":"#microsoft.graph.fileAttachment","id":"a1","name":"one.txt","contentType":"text/plain","size":10,"isInline":false}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/me/messages/m1/attachments?$skiptoken=attachment-next"}
                """),
            Response("""
                {"value":[{"@odata.type":"#microsoft.graph.fileAttachment","id":"a2","name":"two.txt","contentType":"text/plain","size":20,"isInline":false}]}
                """));

        var attachments = await setup.Service.ListAttachmentsAsync("m1");

        attachments.Select(attachment => attachment.Id).Should().Equal("a1", "a2");
        setup.Handler.Requests.Should().HaveCount(2);
        setup.Handler.Requests[1].RequestUri!.Query.Should().Contain("$skiptoken=attachment-next");
    }

    [Fact]
    public async Task ListCategories_follows_next_link()
    {
        var setup = Create(
            Response("""
                {"value":[{"id":"c1","displayName":"One","color":"preset0"}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/me/outlook/masterCategories?$skiptoken=category-next"}
                """),
            Response("""
                {"value":[{"id":"c2","displayName":"Two","color":"preset1"}]}
                """));

        var categories = await setup.Service.ListCategoriesAsync();

        categories.Select(category => category.Id).Should().Equal("c1", "c2");
        setup.Handler.Requests.Should().HaveCount(2);
        setup.Handler.Requests[1].RequestUri!.Query.Should().Contain("$skiptoken=category-next");
    }

    [Fact]
    public async Task Search_with_query_and_sender_uses_separate_escaped_search_clauses()
    {
        var setup = Create(Response("""{"value":[]}"""));

        await setup.Service.SearchAsync(new EmailQuery(
            "subject:\"Q4\"\\path", null, "boss@example.com"));

        string query = Uri.UnescapeDataString(setup.Handler.Requests.Single().RequestUri!.Query);
        query.Should().Contain("$search=\"subject:\\\"Q4\\\"\\\\path\" AND \"from:boss@example.com\"");
    }

    [Fact]
    public async Task Sender_filter_precedes_orderby_and_escapes_o_data_literals()
    {
        var setup = Create(Response("""{"value":[]}"""));

        await setup.Service.SearchAsync(new EmailQuery(
            null, null, "o'hare@example.com"));

        string query = Uri.UnescapeDataString(setup.Handler.Requests.Single().RequestUri!.Query);
        query.Should().Contain(
            "$filter=receivedDateTime ge 1900-01-01T00:00:00Z and from/emailAddress/address eq 'o''hare@example.com'");
        query.Should().Contain("$orderby=receivedDateTime desc");
    }

    [Fact]
    public async Task ReadAttachment_rejects_large_metadata_before_content_request()
    {
        var setup = Create(Response("""
            {"@odata.type":"#microsoft.graph.fileAttachment","id":"a3","name":"big.bin","contentType":"application/octet-stream","size":3000000,"isInline":false}
            """));

        var act = () => setup.Service.ReadAttachmentAsync("m1", "a3", maxBytes: 100);

        await act.Should().ThrowAsync<GraphServiceException>()
            .WithMessage("*[attachment-too-large]*");
        setup.Handler.Requests.Should().ContainSingle();
        Uri.UnescapeDataString(setup.Handler.Requests[0].RequestUri!.Query)
            .Should().NotContain("contentBytes");
    }

    [Fact]
    public async Task ReadAttachment_fetches_content_only_after_metadata_check()
    {
        var setup = Create(
            Response("""
                {"@odata.type":"#microsoft.graph.fileAttachment","id":"a1","name":"one.txt","contentType":"text/plain","size":5,"isInline":false}
                """),
            Response("""
                {"@odata.type":"#microsoft.graph.fileAttachment","id":"a1","name":"one.txt","contentType":"text/plain","size":5,"isInline":false,"contentBytes":"SGVsbG8="}
                """));

        var attachment = await setup.Service.ReadAttachmentAsync("m1", "a1", maxBytes: 100);

        attachment.Text.Should().Be("Hello");
        setup.Handler.Requests.Should().HaveCount(2);
        Uri.UnescapeDataString(setup.Handler.Requests[0].RequestUri!.Query)
            .Should().NotContain("contentBytes");
        Uri.UnescapeDataString(setup.Handler.Requests[1].RequestUri!.Query)
            .Should().Contain("contentBytes");
    }

    [Fact]
    public async Task CreateFolder_with_parent_uses_child_folders_endpoint()
    {
        var setup = Create(Response("""
                {"id":"sophie","displayName":"Sophie","parentFolderId":"kids"}
                """));

        var folder = await setup.Service.CreateFolderAsync("Sophie", "kids");

        folder.Id.Should().Be("sophie");
        folder.Path.Should().BeNull();
        setup.Handler.Requests.Should().ContainSingle();
        setup.Handler.Requests[0].Method.Should().Be(HttpMethod.Post);
        setup.Handler.Requests[0].RequestUri!.AbsolutePath
            .Should().Be("/v1.0/me/mailFolders/kids/childFolders");
    }

    [Fact]
    public async Task Move_resolves_nested_path_to_graph_folder_id()
    {
        var setup = Create(
            Response("""{"value":[{"id":"school","displayName":"School","childFolderCount":1}]}"""),
            Response("""{"value":[{"id":"kids","displayName":"Kids","parentFolderId":"school","childFolderCount":1}]}"""),
            Response("""{"value":[{"id":"sophie","displayName":"Sophie","parentFolderId":"kids","childFolderCount":0}]}"""),
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
        => Create(new GraphAuthOptions(), responses);

    private static (GraphMailService Service, SequenceHandler Handler, HttpClient Client) Create(
        GraphAuthOptions options,
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
            Options.Create(options),
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