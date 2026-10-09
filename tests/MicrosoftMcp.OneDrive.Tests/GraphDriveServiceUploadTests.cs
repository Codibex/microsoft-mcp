using System.Net;
using System.Text;
using AwesomeAssertions;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Kiota.Http.HttpClientLibrary;
using MicrosoftMcp.Common;
using NSubstitute;

namespace MicrosoftMcp.OneDrive.Tests;

/// <summary>Resumable upload against <see cref="GraphDriveService"/> with stubbed
/// HTTP: Graph adapter for drive/session calls, plain client for fragments.</summary>
public sealed class GraphDriveServiceUploadTests
{
    private const int ChunkSize = 5_242_880; // mirrors GraphDriveService.UploadChunkSize

    [Fact]
    public async Task Large_local_file_uses_resumable_session_with_sequential_ranges()
    {
        const int total = 6 * 1024 * 1024;
        string path = Path.GetTempFileName();
        try
        {
            byte[] data = new byte[total];
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = (byte)(i % 251);
            }

            await File.WriteAllBytesAsync(path, data);

            var graphHandler = new SequenceHandler(
                JsonResponse("""{"id":"d1"}"""),
                JsonResponse("""{"id":"root-id"}"""),
                JsonResponse("""{"uploadUrl":"https://upload.test/session","expirationDateTime":"2999-01-01T00:00:00Z"}"""));
            using var harness = new ServiceHarness(graphHandler,
                new HttpResponseMessage(HttpStatusCode.Accepted),
                JsonResponse(
                    """{"id":"up-1","name":"gross.bin","size":6291456,"file":{"mimeType":"application/octet-stream"},"parentReference":{"id":"root-id"},"webUrl":"https://example/gross.bin","lastModifiedDateTime":"2026-10-01T00:00:00Z"}""",
                    HttpStatusCode.OK));

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
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task Small_local_file_uses_simple_put()
    {
        string path = Path.GetTempFileName();
        try
        {
            await File.WriteAllTextAsync(path, "klein");

            var graphHandler = new SequenceHandler(
                JsonResponse("""{"id":"d1"}"""),
                JsonResponse("""{"id":"root-id"}"""),
                JsonResponse("""{"id":"up-2","name":"klein.txt","size":5}""", HttpStatusCode.Created));
            using var harness = new ServiceHarness(graphHandler);

            var summary = await harness.Service.UploadAsync("klein.txt", "root", localPath: path);

            summary.Id.Should().Be("up-2");
            graphHandler.Requests.Should().HaveCount(3);
            graphHandler.Requests[2].Method.Should().Be(HttpMethod.Put);
            graphHandler.Requests[2].RequestUri!.ToString().Should().Contain("content");
            harness.UploadHandler.Requests.Should().BeEmpty();
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task Upload_rejects_relative_and_missing_paths()
    {
        using var harness = new ServiceHarness(new SequenceHandler());
        var service = harness.Service;

        await Assert.ThrowsAsync<GraphServiceException>(() =>
            service.UploadAsync("x.txt", "root", localPath: "relativ/datei.txt"));
        await Assert.ThrowsAsync<GraphServiceException>(() =>
            service.UploadAsync("x.txt", "root",
                localPath: Path.Combine(Path.GetTempPath(), Path.GetRandomFileName())));
    }

    private sealed class ServiceHarness : IDisposable
    {
        private readonly HttpClient _graphHttp;
        private readonly HttpClient _uploadHttp;

        public GraphDriveService Service { get; }
        public SequenceHandler UploadHandler { get; } = new();

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
                _uploadHttp);
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
