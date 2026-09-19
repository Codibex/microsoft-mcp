using System.Net;
using System.Text;
using AwesomeAssertions;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Kiota.Http.HttpClientLibrary;
using MicrosoftMcp.Common;
using NSubstitute;

namespace MicrosoftMcp.Teams.Tests;

public sealed class GraphTeamsServiceTests
{
    [Fact]
    public async Task ListTeams_does_not_send_unsupported_query_parameters()
    {
        var handler = new RecordingHandler();
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var graphClient = new GraphServiceClient(requestAdapter);
        var service = new GraphTeamsService(
            graphClient,
            Options.Create(new GraphAuthOptions()));

        var teams = await service.ListTeamsAsync();

        teams.Should().ContainSingle(team => team.Id == "team-1");
        handler.RequestUri.Should().NotBeNull();
        handler.RequestUri!.AbsolutePath.Should().Be("/v1.0/me/joinedTeams");
        handler.RequestUri.Query.Should().BeEmpty();
    }

    private sealed class RecordingHandler : HttpMessageHandler
    {
        public Uri? RequestUri { get; private set; }

        protected override Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request,
            CancellationToken cancellationToken)
        {
            RequestUri = request.RequestUri;
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    "{\"value\":[{\"id\":\"team-1\",\"displayName\":\"Engineering\"}]}",
                    Encoding.UTF8,
                    "application/json")
            });
        }
    }
}