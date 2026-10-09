using System.Net;
using System.Text;
using AwesomeAssertions;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Kiota.Http.HttpClientLibrary;
using MicrosoftMcp.Common;
using NSubstitute;
using System.IO.Abstractions.TestingHelpers;

namespace MicrosoftMcp.SharePoint.Tests;

/// <summary>Sites discovery, explicit-drive transfers and error mapping
/// against <see cref="GraphSharePointService"/> with stubbed HTTP and an
/// in-memory filesystem: no Graph, no network, no disk.</summary>
public sealed class GraphSharePointServiceTransferTests
{
    private const int ChunkSize = 5_242_880; // mirrors GraphSharePointService.UploadChunkSize

    private static string TempPath(string name) => Path.Combine(Path.GetTempPath(), name);

    [Fact]
    public async Task Search_sites_uses_supported_path_and_query()
    {
        var graphHandler = new SequenceHandler(
            JsonResponse("""{"value":[{"id":"site-eng","name":"Engineering","displayName":"Engineering","webUrl":"https://contoso.sharepoint.com/sites/eng"}]}"""));
        using var harness = new ServiceHarness(graphHandler);

        var sites = await harness.Service.SearchSitesAsync("Engineering");

        sites.Should().ContainSingle(s => s.Id == "site-eng");
        graphHandler.Requests.Should().HaveCount(1);
        graphHandler.Requests[0].RequestUri!.AbsolutePath.Should().Be("/v1.0/sites");
        graphHandler.Requests[0].RequestUri!.Query.Should().Contain("search=Engineering");
    }

    [Fact]
    public async Task List_site_drives_uses_supported_path()
    {
        var graphHandler = new SequenceHandler(
            JsonResponse("""{"value":[{"id":"drive-eng","name":"Dokumente","driveType":"documentLibrary","quota":{"total":10000,"used":1000,"remaining":9000}}]}"""));
        using var harness = new ServiceHarness(graphHandler);

        var drives = await harness.Service.ListSiteDrivesAsync("site-eng");

        drives.Should().ContainSingle(d => d.Id == "drive-eng");
        graphHandler.Requests.Should().HaveCount(1);
        graphHandler.Requests[0].RequestUri!.AbsolutePath.Should().Be("/v1.0/sites/site-eng/drives");
    }

    [Fact]
    public async Task Unknown_drive_reports_drive_not_found_not_item_not_found()
    {
        var graphHandler = new SequenceHandler(
            JsonResponse("""{"error":{"code":"itemNotFound","message":"drive not found"}}""", HttpStatusCode.NotFound));
        using var harness = new ServiceHarness(graphHandler);

        var ex = await Assert.ThrowsAsync<GraphServiceException>(
            () => harness.Service.GetItemAsync("drive-nope", "root"));

        ex.Code.Should().Be("drive-not-found");
        ex.Message.Should().Contain("sharepoint_list_site_drives");
    }

    [Fact]
    public async Task List_children_pages_and_sorts_folders_first()
    {
        var graphHandler = new SequenceHandler(
            JsonResponse("""{"id":"root-id"}"""),
            JsonResponse("""{"value":[{"id":"f-file","name":"b.txt","size":3,"file":{"mimeType":"text/plain"}}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/drives/drive-eng/items/root-id/children?$skiptoken=next"}"""),
            JsonResponse("""{"value":[{"id":"f-folder","name":"a-folder","folder":{"childCount":1}}]}"""));
        using var harness = new ServiceHarness(graphHandler);

        var children = await harness.Service.ListChildrenAsync("drive-eng", "root");

        children.Select(i => i.Id).Should().Equal("f-folder", "f-file");
        graphHandler.Requests.Should().HaveCount(3);
        graphHandler.Requests[0].RequestUri!.AbsolutePath.Should().Be("/v1.0/drives/drive-eng/root");
        graphHandler.Requests[2].RequestUri!.Query.Should().Contain("$skiptoken=next");
    }

    [Fact]
    public async Task Small_local_file_uses_simple_put_on_explicit_drive()
    {
        var graphHandler = new SequenceHandler(
            JsonResponse("""{"id":"root-id"}"""),
            JsonResponse("""{"value":[]}"""),
            JsonResponse("""{"id":"up-1","name":"klein.txt","size":5}""", HttpStatusCode.Created));
        using var harness = new ServiceHarness(graphHandler);

        string path = TempPath("sp-klein.txt");
        harness.FileSystem.AddFile(path, "klein");

        var summary = await harness.Service.UploadAsync("drive-eng", "klein.txt", "root", localPath: path);

        summary.Id.Should().Be("up-1");
        graphHandler.Requests.Should().HaveCount(3);
        graphHandler.Requests[2].Method.Should().Be(HttpMethod.Put);
        graphHandler.Requests[2].RequestUri!.AbsolutePath.Should().Contain("/v1.0/drives/drive-eng/");
        graphHandler.Requests[2].RequestUri!.ToString().Should().Contain("content");
        harness.UploadHandler.Requests.Should().BeEmpty();
    }

    [Fact]
    public async Task Large_local_file_uses_resumable_session_with_sequential_ranges()
    {
        const int total = 6 * 1024 * 1024;
        var graphHandler = new SequenceHandler(
            JsonResponse("""{"id":"root-id"}"""),
            JsonResponse("""{"uploadUrl":"https://upload.test/session","expirationDateTime":"2999-01-01T00:00:00Z"}"""));
        using var harness = new ServiceHarness(graphHandler,
            new HttpResponseMessage(HttpStatusCode.Accepted),
            JsonResponse(
                """{"id":"up-2","name":"gross.bin","size":6291456,"file":{"mimeType":"application/octet-stream"},"parentReference":{"id":"root-id"},"webUrl":"https://example/gross.bin","lastModifiedDateTime":"2026-10-01T00:00:00Z"}""",
                HttpStatusCode.OK));

        string path = TempPath("sp-gross.bin");
        harness.FileSystem.AddFile(path, new MockFileData(new byte[total]));

        var summary = await harness.Service.UploadAsync("drive-eng", "gross.bin", "root", localPath: path);

        summary.Id.Should().Be("up-2");
        summary.Size.Should().Be(total);
        graphHandler.Requests.Should().HaveCount(2);
        graphHandler.Requests[1].Method.Should().Be(HttpMethod.Post);
        graphHandler.Requests[1].RequestUri!.AbsolutePath.Should().Contain("createUploadSession");

        var uploadHandler = harness.UploadHandler;
        uploadHandler.Requests.Should().HaveCount(2);
        uploadHandler.Requests[0].Headers.Authorization.Should().BeNull();
        uploadHandler.Requests[0].Content!.Headers.ContentRange!.ToString()
            .Should().Be($"bytes 0-{ChunkSize - 1}/{total}");
        uploadHandler.Requests[1].Content!.Headers.ContentRange!.ToString()
            .Should().Be($"bytes {ChunkSize}-{total - 1}/{total}");
    }

    [Fact]
    public async Task Download_enforces_cap_on_the_stream_not_on_metadata()
    {
        // Stale metadata claims 10 bytes, the stream carries 100: the bounded
        // copy must reject at maxBytes + 1 without trusting size.
        var graphHandler = new SequenceHandler(
            JsonResponse("""{"id":"root-id"}"""),
            JsonResponse("""{"id":"f-big","name":"gross.bin","size":10,"file":{"mimeType":"application/octet-stream"}}"""),
            new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(new string('x', 100), Encoding.UTF8, "application/octet-stream")
            });
        using var harness = new ServiceHarness(graphHandler);

        var ex = await Assert.ThrowsAsync<GraphServiceException>(
            () => harness.Service.DownloadAsync("drive-eng", "f-big", maxBytes: 50));

        ex.Code.Should().Be("attachment-too-large");
    }

    private sealed class ServiceHarness : IDisposable
    {
        private readonly HttpClient _graphHttp;
        private readonly HttpClient _uploadHttp;

        public GraphSharePointService Service { get; }
        public SequenceHandler UploadHandler { get; } = new();
        public MockFileSystem FileSystem { get; } = new();

        public ServiceHarness(SequenceHandler graphHandler, params HttpResponseMessage[] uploadResponses)
        {
            foreach (var response in uploadResponses)
            {
                UploadHandler.Enqueue(response);
            }

            _graphHttp = new HttpClient(graphHandler);
            var adapter = new HttpClientRequestAdapter(
                Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
                httpClient: _graphHttp);
            _uploadHttp = new HttpClient(UploadHandler);
            Service = new GraphSharePointService(
                new GraphServiceClient(adapter),
                Options.Create(new GraphAuthOptions()),
                _uploadHttp,
                FileSystem);
        }

        public void Dispose()
        {
            _graphHttp.Dispose();
            _uploadHttp.Dispose();
        }
    }

    private static HttpResponseMessage JsonResponse(string json, HttpStatusCode status = HttpStatusCode.OK) =>
        new(status) { Content = new StringContent(json, Encoding.UTF8, "application/json") };

    private sealed class SequenceHandler : HttpMessageHandler
    {
        private readonly Queue<HttpResponseMessage> _responses = new();

        public List<HttpRequestMessage> Requests { get; } = [];

        public SequenceHandler(params HttpResponseMessage[] responses)
        {
            Enqueue(responses);
        }

        public void Enqueue(params HttpResponseMessage[] responses)
        {
            foreach (var response in responses)
            {
                _responses.Enqueue(response);
            }
        }

        protected override Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request, CancellationToken cancellationToken)
        {
            Requests.Add(request);
            return Task.FromResult(_responses.Dequeue());
        }
    }
}
