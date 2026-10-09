using System.Net;
using System.Text;
using AwesomeAssertions;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Kiota.Http.HttpClientLibrary;
using MicrosoftMcp.Common;
using NSubstitute;
using System.IO.Abstractions.TestingHelpers;

namespace MicrosoftMcp.OneDrive.Tests;

/// <summary>Upload sessions and file downloads against <see cref="GraphDriveService"/>
/// with stubbed HTTP and an in-memory filesystem: no Graph, no network, no disk.</summary>
public sealed class GraphDriveServiceTransferTests
{
    private const int ChunkSize = 5_242_880; // mirrors GraphDriveService.UploadChunkSize

    private static string TempPath(string name) => Path.Combine(Path.GetTempPath(), name);

    [Fact]
    public async Task Large_local_file_uses_resumable_session_with_sequential_ranges()
    {
        const int total = 6 * 1024 * 1024;
        var graphHandler = new SequenceHandler(
            JsonResponse("""{"id":"d1"}"""),
            JsonResponse("""{"id":"root-id"}"""),
            JsonResponse("""{"uploadUrl":"https://upload.test/session","expirationDateTime":"2999-01-01T00:00:00Z"}"""));
        using var harness = new ServiceHarness(graphHandler,
            new HttpResponseMessage(HttpStatusCode.Accepted),
            JsonResponse(
                """{"id":"up-1","name":"gross.bin","size":6291456,"file":{"mimeType":"application/octet-stream"},"parentReference":{"id":"root-id"},"webUrl":"https://example/gross.bin","lastModifiedDateTime":"2026-10-01T00:00:00Z"}""",
                HttpStatusCode.OK));

        string path = TempPath("gross.bin");
        harness.FileSystem.AddFile(path, new MockFileData(new byte[total]));

        var summary = await harness.Service.UploadAsync("gross.bin", "root", localPath: path);

        summary.Id.Should().Be("up-1");
        summary.Name.Should().Be("gross.bin");
        summary.Size.Should().Be(total);
        summary.IsFolder.Should().BeFalse();
        summary.ParentId.Should().Be("root-id");

        graphHandler.Requests.Should().HaveCount(3);
        graphHandler.Requests[2].Method.Should().Be(HttpMethod.Post);
        graphHandler.Requests[2].RequestUri!.AbsolutePath.Should().Contain("createUploadSession");

        var uploadHandler = harness.UploadHandler;
        uploadHandler.Requests.Should().HaveCount(2);
        uploadHandler.Requests[0].Headers.Authorization.Should().BeNull();
        uploadHandler.Requests[0].Content!.Headers.ContentRange!.ToString()
            .Should().Be($"bytes 0-{ChunkSize - 1}/{total}");
        uploadHandler.Requests[1].Content!.Headers.ContentRange!.ToString()
            .Should().Be($"bytes {ChunkSize}-{total - 1}/{total}");
    }

    [Fact]
    public async Task Small_local_file_uses_simple_put()
    {
        var graphHandler = new SequenceHandler(
            JsonResponse("""{"id":"d1"}"""),
            JsonResponse("""{"id":"root-id"}"""),
            JsonResponse("""{"value":[]}"""),
            JsonResponse("""{"id":"up-2","name":"klein.txt","size":5}""", HttpStatusCode.Created));
        using var harness = new ServiceHarness(graphHandler);

        string path = TempPath("klein.txt");
        harness.FileSystem.AddFile(path, "klein");

        var summary = await harness.Service.UploadAsync("klein.txt", "root", localPath: path);

        summary.Id.Should().Be("up-2");
        graphHandler.Requests.Should().HaveCount(4);
        graphHandler.Requests[3].Method.Should().Be(HttpMethod.Put);
        graphHandler.Requests[3].RequestUri!.ToString().Should().Contain("content");
        harness.UploadHandler.Requests.Should().BeEmpty();
    }

    [Fact]
    public async Task Simple_upload_renames_on_name_conflict()
    {
        var graphHandler = new SequenceHandler(
            JsonResponse("""{"id":"d1"}"""),
            JsonResponse("""{"id":"root-id"}"""),
            JsonResponse("""{"value":[{"id":"old","name":"dup.txt"}]}"""),
            JsonResponse("""{"id":"up-3","name":"dup 1.txt","size":5}""", HttpStatusCode.Created));
        using var harness = new ServiceHarness(graphHandler);

        string path = TempPath("dup-src.txt");
        harness.FileSystem.AddFile(path, "neu");

        var summary = await harness.Service.UploadAsync("dup.txt", "root", localPath: path);

        summary.Name.Should().Be("dup 1.txt");
        string putUrl = graphHandler.Requests[3].RequestUri!.ToString();
        putUrl.Should().Contain("dup").And.NotContain("dup.txt");
    }

    [Fact]
    public async Task Session_upload_applies_server_side_rename()
    {
        const int total = 6 * 1024 * 1024;
        var graphHandler = new SequenceHandler(
            JsonResponse("""{"id":"d1"}"""),
            JsonResponse("""{"id":"root-id"}"""),
            JsonResponse("""{"uploadUrl":"https://upload.test/session","expirationDateTime":"2999-01-01T00:00:00Z"}"""));
        using var harness = new ServiceHarness(graphHandler,
            new HttpResponseMessage(HttpStatusCode.Accepted),
            JsonResponse("""{"id":"up-4","name":"big 1.bin","size":6291456}""", HttpStatusCode.OK));

        string path = TempPath("big.bin");
        harness.FileSystem.AddFile(path, new MockFileData(new byte[total]));

        var summary = await harness.Service.UploadAsync("big.bin", "root", localPath: path);

        summary.Name.Should().Be("big 1.bin");
    }

    [Fact]
    public async Task Upload_rejects_relative_and_missing_paths()
    {
        using var harness = new ServiceHarness(new SequenceHandler());
        var service = harness.Service;

        (await Assert.ThrowsAsync<GraphServiceException>(() =>
            service.UploadAsync("x.txt", "root", localPath: "relativ/datei.txt")))
            .Code.Should().Be("invalid-request");
        (await Assert.ThrowsAsync<GraphServiceException>(() =>
            service.UploadAsync("x.txt", "root", localPath: TempPath(Path.GetRandomFileName()))))
            .Code.Should().Be("invalid-request");
    }

    [Fact]
    public async Task Download_to_file_streams_bytes_to_disk()
    {
        byte[] payload = Encoding.UTF8.GetBytes("Hallo von Graph");
        var graphHandler = new SequenceHandler(
            JsonResponse("""{"id":"d1"}"""),
            JsonResponse("""{"id":"root-id"}"""),
            JsonResponse("""{"id":"f-1","name":"a.txt","size":15,"file":{"mimeType":"text/plain"},"parentReference":{"id":"root-id"}}"""),
            new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(payload) });
        using var harness = new ServiceHarness(graphHandler);

        string dest = TempPath("dl-a.txt");
        var result = await harness.Service.DownloadToFileAsync("f-1", dest);

        result.Encoding.Should().Be("file");
        result.LocalPath.Should().Be(dest);
        result.Size.Should().Be(payload.Length);
        result.Text.Should().BeNull();
        result.DataBase64.Should().BeNull();
        harness.FileSystem.File.ReadAllBytes(dest).Should().Equal(payload);
        graphHandler.Requests.Should().HaveCount(4);
        graphHandler.Requests[3].RequestUri!.ToString().Should().Contain("content");
    }

    [Fact]
    public async Task Download_to_file_refuses_overwrite_before_any_network_call()
    {
        var graphHandler = new SequenceHandler();
        using var harness = new ServiceHarness(graphHandler);

        string dest = TempPath("dl-bestehend.txt");
        harness.FileSystem.AddFile(dest, "alt");

        (await Assert.ThrowsAsync<GraphServiceException>(() =>
            harness.Service.DownloadToFileAsync("f-1", dest)))
            .Code.Should().Be("invalid-request");
        graphHandler.Requests.Should().BeEmpty();
        harness.FileSystem.File.ReadAllText(dest).Should().Be("alt");
    }

    [Fact]
    public async Task Failed_download_leaves_no_partial_file()
    {
        var graphHandler = new SequenceHandler(
            JsonResponse("""{"id":"d1"}"""),
            JsonResponse("""{"id":"root-id"}"""),
            JsonResponse("""{"id":"f-1","name":"a.txt","size":100,"file":{"mimeType":"application/octet-stream"}}"""),
            new HttpResponseMessage(HttpStatusCode.OK) { Content = new FailingContent() });
        using var harness = new ServiceHarness(graphHandler);

        string dest = TempPath("dl-defekt.bin");
        var ex = await Assert.ThrowsAsync<HttpRequestException>(() =>
            harness.Service.DownloadToFileAsync("f-1", dest));
        ex.InnerException.Should().BeOfType<IOException>();

        harness.FileSystem.File.Exists(dest).Should().BeFalse();
        harness.FileSystem.Directory.GetFiles(harness.FileSystem.Path.GetDirectoryName(dest)!)
            .Should().BeEmpty();
    }

    private sealed class ServiceHarness : IDisposable
    {
        private readonly HttpClient _graphHttp;
        private readonly HttpClient _uploadHttp;

        public GraphDriveService Service { get; }
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
            Service = new GraphDriveService(
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

    /// <summary>Response body that fails mid-stream to simulate a broken download.</summary>
    private sealed class FailingContent : HttpContent
    {
        protected override async Task SerializeToStreamAsync(Stream stream, TransportContext? context)
        {
            await stream.WriteAsync(new byte[3]);
            throw new IOException("boom");
        }

        protected override bool TryComputeLength(out long length)
        {
            length = 100;
            return true;
        }
    }
}
