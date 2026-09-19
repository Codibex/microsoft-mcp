using System.Net;
using System.Text;
using AwesomeAssertions;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Kiota.Http.HttpClientLibrary;
using MicrosoftMcp.Common;
using NSubstitute;
using System.Text.Json;

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

    [Fact]
    public async Task ListChannels_follows_graph_paging_without_adding_top()
    {
        var handler = new SequenceHandler(
            Response("""
                {"value":[{"id":"channel-1","displayName":"General"}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/teams/team-1/channels?$skiptoken=next"}
                """),
            Response("""
                {"value":[{"id":"channel-2","displayName":"Planning"}]}
                """));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var service = new GraphTeamsService(
            new GraphServiceClient(requestAdapter),
            Options.Create(new GraphAuthOptions()));

        var channels = await service.ListChannelsAsync("team-1");

        channels.Select(channel => channel.Id).Should().Equal("channel-1", "channel-2");
        handler.Requests.Should().HaveCount(2);
        handler.Requests[0].RequestUri!.AbsolutePath.Should().Be("/v1.0/teams/team-1/channels");
        handler.Requests[0].RequestUri!.Query.Should().NotContain("$top");
        handler.Requests[1].RequestUri!.Query.Should().Contain("$skiptoken=next");
    }

    [Fact]
    public async Task ReadOnly_message_and_chat_reads_use_supported_paths_and_queries()
    {
        var handler = new SequenceHandler(
            Response("""{"value":[]}"""),
            Response("""{"value":[]}"""),
            Response("""{"value":[]}"""),
            Response("""{"id":"message-1","messageType":"message","body":{"contentType":"html","content":"Channel"}}"""),
            Response("""{"value":[]}"""),
            Response("""{"value":[]}"""),
            Response("""{"id":"message-1","messageType":"message","body":{"contentType":"html","content":"Chat"}}"""));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var service = new GraphTeamsService(
            new GraphServiceClient(requestAdapter),
            Options.Create(new GraphAuthOptions()));

        await service.ListChannelsAsync("team-1");
        await service.ListChannelMessagesAsync("team-1", "channel-1", 100);
        await service.ListMessageRepliesAsync("team-1", "channel-1", "message-1", 100);
        await service.ReadChannelMessageAsync("team-1", "channel-1", "message-1");
        await service.ListChatsAsync(100);
        await service.ListChatMessagesAsync("19:3d214d1a-1234-4567-8901-abcdef123456@thread.v2", 100);
        await service.ReadChatMessageAsync("chat-1", "message-1");

        handler.Requests.Should().HaveCount(7);
        handler.Requests[0].RequestUri!.AbsolutePath.Should().Be("/v1.0/teams/team-1/channels");
        handler.Requests[0].RequestUri!.Query.Should().NotContain("$top");
        handler.Requests[1].RequestUri!.AbsolutePath.Should().Be("/v1.0/teams/team-1/channels/channel-1/messages");
        handler.Requests[2].RequestUri!.AbsolutePath.Should().Be("/v1.0/teams/team-1/channels/channel-1/messages/message-1/replies");
        handler.Requests[4].RequestUri!.AbsolutePath.Should().Be("/v1.0/me/chats");
        Uri.UnescapeDataString(handler.Requests[5].RequestUri!.AbsolutePath).Should().Be(
            "/v1.0/me/chats/19:3d214d1a-1234-4567-8901-abcdef123456@thread.v2/messages");
        handler.Requests[6].RequestUri!.AbsolutePath.Should().Be("/v1.0/me/chats/chat-1/messages/message-1");
        foreach (var request in handler.Requests)
        {
            request.RequestUri!.Query.Should().NotContain("$select");
        }

        Uri.UnescapeDataString(handler.Requests[1].RequestUri!.Query).Should().Contain("$top=50");
        Uri.UnescapeDataString(handler.Requests[2].RequestUri!.Query).Should().Contain("$top=50");
        Uri.UnescapeDataString(handler.Requests[4].RequestUri!.Query).Should().Contain("$top=50");
        Uri.UnescapeDataString(handler.Requests[5].RequestUri!.Query).Should().Contain("$top=50");
        handler.Requests[3].RequestUri!.Query.Should().BeEmpty();
        handler.Requests[6].RequestUri!.Query.Should().BeEmpty();
    }

    [Fact]
    public async Task ListMeetingInsights_uses_documented_path_without_top_and_follows_paging()
    {
        var handler = new SequenceHandler(
            Response("""
                {"value":[{"id":"insight-1","callId":"call-1"}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/copilot/users/user-1/onlineMeetings/meeting-1/aiInsights?$skiptoken=next"}
                """),
            Response("""{"value":[{"id":"insight-2","callId":"call-2"}]}"""));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var service = new GraphTeamsService(
            new GraphServiceClient(requestAdapter),
            Options.Create(new GraphAuthOptions { UserIdOrUpn = "user-1" }));

        var insights = await service.ListMeetingInsightsAsync("meeting-1", 100);

        insights.Select(insight => insight.Id).Should().Equal("insight-1", "insight-2");
        handler.Requests.Should().HaveCount(2);
        handler.Requests[0].RequestUri!.AbsolutePath.Should().Be(
            "/v1.0/copilot/users/user-1/onlineMeetings/meeting-1/aiInsights");
        handler.Requests[0].RequestUri!.Query.Should().BeEmpty();
        handler.Requests[1].RequestUri!.Query.Should().Contain("$skiptoken=next");
    }

    [Fact]
    public async Task SendChannelMessage_checks_members_and_applies_disclosure_before_post()
    {
        var handler = new SequenceHandler(
            Response("""
                {"value":[{"@odata.type":"#microsoft.graph.aadUserConversationMember","id":"member-1","displayName":"Alice","email":"alice@firma.de","tenantId":"tenant-1","userId":"user-1"}]}
                """),
            Response("""
                {"id":"message-1","messageType":"message","body":{"contentType":"html","content":"<p>Sent</p>"}}
                """));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var graphClient = new GraphServiceClient(requestAdapter);
        var policy = new TeamsPolicyOptions
        {
            RequireInternalRecipients = true,
            AllowedRecipientDomains = ["firma.de"],
            AiDisclosureEnabled = true,
            AiDisclosureText = "KI-Hinweis"
        };
        var service = new GraphTeamsService(
            graphClient,
            Options.Create(new GraphAuthOptions { TenantId = "tenant-1" }),
            Options.Create(policy));

        MessageDetail result = await service.SendChannelMessageAsync("team-1", "channel-1", "<p>Hello</p>");

        result.Id.Should().Be("message-1");
        handler.Requests.Should().HaveCount(2);
        handler.Requests[0].Method.Should().Be(HttpMethod.Get);
        handler.Requests[0].RequestUri!.AbsolutePath.Should().Be("/v1.0/teams/team-1/channels/channel-1/allMembers");
        handler.Requests[1].Method.Should().Be(HttpMethod.Post);
        handler.Requests[1].RequestUri!.AbsolutePath.Should().Be("/v1.0/teams/team-1/channels/channel-1/messages");
        using JsonDocument payload = JsonDocument.Parse(handler.RequestBodies[1]);
        string content = payload.RootElement.GetProperty("body").GetProperty("content").GetString()!;
        content.Should().Contain("<p>Hello</p>");
        content.Should().Contain("KI-Hinweis");
    }

    [Fact]
    public async Task SendChannelMessage_rejects_external_member_before_post()
    {
        var handler = new SequenceHandler(
            Response("""
                {"value":[{"@odata.type":"#microsoft.graph.aadUserConversationMember","id":"member-1","displayName":"Guest","email":"guest@external.example","tenantId":"foreign-tenant","userId":"user-1"}]}
                """),
            Response("""{"id":"must-not-be-sent"}"""));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var graphClient = new GraphServiceClient(requestAdapter);
        var service = new GraphTeamsService(
            graphClient,
            Options.Create(new GraphAuthOptions { TenantId = "tenant-1" }),
            Options.Create(new TeamsPolicyOptions
            {
                RequireInternalRecipients = true,
                AllowedRecipientDomains = ["firma.de"]
            }));

        var exception = await Assert.ThrowsAsync<GraphServiceException>(
            () => service.SendChannelMessageAsync("team-1", "channel-1", "Hello"));

        exception.Code.Should().Be("invalid-request");
        handler.Requests.Should().ContainSingle();
    }

    [Fact]
    public async Task SendChannelMessage_checks_all_member_pages_before_post()
    {
        var handler = new SequenceHandler(
            Response("""
                {"value":[{"@odata.type":"#microsoft.graph.aadUserConversationMember","id":"member-1","displayName":"Alice","email":"alice@firma.de","tenantId":"tenant-1","userId":"user-1"}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/teams/team-1/channels/channel-1/allMembers?$skiptoken=next"}
                """),
            Response("""
                {"value":[{"@odata.type":"#microsoft.graph.aadUserConversationMember","id":"member-2","displayName":"Guest","email":"guest@external.example","tenantId":"foreign-tenant","userId":"user-2"}]}
                """),
            Response("""{"id":"must-not-be-sent"}"""));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var graphClient = new GraphServiceClient(requestAdapter);
        var service = new GraphTeamsService(
            graphClient,
            Options.Create(new GraphAuthOptions { TenantId = "tenant-1" }),
            Options.Create(new TeamsPolicyOptions
            {
                RequireInternalRecipients = true,
                AllowedRecipientDomains = ["firma.de"]
            }));

        var exception = await Assert.ThrowsAsync<GraphServiceException>(
            () => service.SendChannelMessageAsync("team-1", "channel-1", "Hello"));

        exception.Code.Should().Be("invalid-request");
        handler.Requests.Should().HaveCount(2);
        handler.Requests[1].RequestUri!.Query.Should().Contain("$skiptoken=next");
    }

    [Fact]
    public async Task SendChannelMessage_rejects_missing_tenant_before_post()
    {
        var handler = new SequenceHandler(
            Response("""
                {"value":[{"@odata.type":"#microsoft.graph.aadUserConversationMember","id":"member-1","displayName":"Alice","email":"alice@firma.de","userId":"user-1"}]}
                """),
            Response("""{"id":"must-not-be-sent"}"""));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var graphClient = new GraphServiceClient(requestAdapter);
        var service = new GraphTeamsService(
            graphClient,
            Options.Create(new GraphAuthOptions { TenantId = "tenant-1" }),
            Options.Create(new TeamsPolicyOptions
            {
                RequireInternalRecipients = true,
                AllowedRecipientDomains = ["firma.de"]
            }));

        var exception = await Assert.ThrowsAsync<GraphServiceException>(
            () => service.SendChannelMessageAsync("team-1", "channel-1", "Hello"));

        exception.Code.Should().Be("invalid-request");
        handler.Requests.Should().ContainSingle();
    }

    [Fact]
    public async Task SendChannelMessage_rejects_guest_role_before_post()
    {
        var handler = new SequenceHandler(
            Response("""
                {"value":[{"@odata.type":"#microsoft.graph.aadUserConversationMember","id":"member-1","displayName":"Guest","email":"guest@firma.de","tenantId":"tenant-1","userId":"user-1","roles":["guest"]}]}
                """),
            Response("""{"id":"must-not-be-sent"}"""));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var graphClient = new GraphServiceClient(requestAdapter);
        var service = new GraphTeamsService(
            graphClient,
            Options.Create(new GraphAuthOptions { TenantId = "tenant-1" }),
            Options.Create(new TeamsPolicyOptions
            {
                RequireInternalRecipients = true,
                AllowedRecipientDomains = ["firma.de"]
            }));

        var exception = await Assert.ThrowsAsync<GraphServiceException>(
            () => service.SendChannelMessageAsync("team-1", "channel-1", "Hello"));

        exception.Code.Should().Be("invalid-request");
        handler.Requests.Should().ContainSingle();
    }

    [Fact]
    public async Task SendChatMessage_posts_to_the_existing_chat_endpoint()
    {
        var handler = new SequenceHandler(
            Response("""
                {"id":"message-1","messageType":"message","body":{"contentType":"html","content":"Hello"}}
                """));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var graphClient = new GraphServiceClient(requestAdapter);
        var service = new GraphTeamsService(
            graphClient,
            Options.Create(new GraphAuthOptions()));

        await service.SendChatMessageAsync("chat-1", "Hello");

        handler.Requests.Should().ContainSingle();
        handler.Requests[0].Method.Should().Be(HttpMethod.Post);
        handler.Requests[0].RequestUri!.AbsolutePath.Should().Be("/v1.0/chats/chat-1/messages");
    }

    [Fact]
    public async Task SendChatMessage_checks_all_member_pages_before_post()
    {
        var handler = new SequenceHandler(
            Response("""
                {"value":[{"@odata.type":"#microsoft.graph.aadUserConversationMember","id":"member-1","displayName":"Alice","email":"alice@firma.de","tenantId":"tenant-1","userId":"user-1"}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/me/chats/chat-1/members?$skiptoken=next"}
                """),
            Response("""
                {"value":[{"@odata.type":"#microsoft.graph.aadUserConversationMember","id":"member-2","displayName":"Guest","email":"guest@firma.de","tenantId":"tenant-1","userId":"user-2","roles":["guest"]}]}
                """),
            Response("""{"id":"must-not-be-sent"}"""));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var graphClient = new GraphServiceClient(requestAdapter);
        var service = new GraphTeamsService(
            graphClient,
            Options.Create(new GraphAuthOptions { TenantId = "tenant-1" }),
            Options.Create(new TeamsPolicyOptions
            {
                RequireInternalRecipients = true,
                AllowedRecipientDomains = ["firma.de"]
            }));

        var exception = await Assert.ThrowsAsync<GraphServiceException>(
            () => service.SendChatMessageAsync("chat-1", "Hello"));

        exception.Code.Should().Be("invalid-request");
        handler.Requests.Should().HaveCount(2);
        handler.Requests[1].RequestUri!.Query.Should().Contain("$skiptoken=next");
    }

    private static HttpResponseMessage Response(string json) => new(HttpStatusCode.OK)
    {
        Content = new StringContent(json, Encoding.UTF8, "application/json")
    };

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