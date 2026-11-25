using System.Net;
using System.Text;
using System.Text.Json;
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
    public async Task ListChats_follows_paging_expands_members_and_sorts_by_last_activity()
    {
        var handler = new SequenceHandler(
            Response("""
                {"value":[{"id":"chat-old","chatType":"oneOnOne","lastUpdatedDateTime":"2025-01-01T00:00:00.000Z"}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/me/chats?$skiptoken=next"}
                """),
            Response("""
                {"value":[{"id":"chat-new","chatType":"oneOnOne","lastUpdatedDateTime":"2026-11-25T09:00:00.000Z","members":[{"@odata.type":"#microsoft.graph.aadUserConversationMember","id":"m-1","displayName":"Bob","email":"bob@contoso.com","userId":"u-bob"}]}]}
                """));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var service = new GraphTeamsService(
            new GraphServiceClient(requestAdapter),
            Options.Create(new GraphAuthOptions()));

        var chats = await service.ListChatsAsync();

        // Graph pages in an unspecified order: the fresher chat from the
        // second page must win over the older one from the first page.
        chats.Select(chat => chat.Id).Should().Equal("chat-new", "chat-old");
        chats[0].Members.Should().ContainSingle(m =>
            m.DisplayName == "Bob" && m.Email == "bob@contoso.com" && m.UserId == "u-bob");
        chats[1].Members.Should().BeEmpty();
        handler.Requests.Should().HaveCount(2);
        Uri.UnescapeDataString(handler.Requests[0].RequestUri!.Query).Should().Contain("$top=50");
        Uri.UnescapeDataString(handler.Requests[0].RequestUri!.Query).Should().Contain("$expand=members");
        handler.Requests[1].RequestUri!.Query.Should().Contain("$skiptoken=next");
    }

    [Fact]
    public async Task ListChats_falls_back_to_plain_listing_when_members_expand_is_rejected()
    {
        var handler = new SequenceHandler(
            new HttpResponseMessage(HttpStatusCode.BadRequest)
            {
                Content = new StringContent(
                    """{"error":{"code":"Request_BadRequest","message":"expand is not supported"}}""",
                    Encoding.UTF8,
                    "application/json")
            },
            Response("""
                {"value":[{"id":"chat-1","chatType":"oneOnOne","lastUpdatedDateTime":"2026-11-25T09:00:00.000Z"}]}
                """));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var service = new GraphTeamsService(
            new GraphServiceClient(requestAdapter),
            Options.Create(new GraphAuthOptions()));

        var chats = await service.ListChatsAsync();

        chats.Should().ContainSingle(chat => chat.Id == "chat-1");
        chats[0].Members.Should().BeEmpty();
        handler.Requests.Should().HaveCount(2);
        Uri.UnescapeDataString(handler.Requests[0].RequestUri!.Query).Should().Contain("$expand=members");
        Uri.UnescapeDataString(handler.Requests[1].RequestUri!.Query).Should().NotContain("$expand");
    }

    [Fact]
    public async Task ListChats_stops_at_the_scan_cap_without_fetching_a_discarded_page()
    {
        // 10 pages × $top=50 = the 500-chat budget; every page advertises a
        // continuation, so an 11th request would be fetched and thrown away.
        var handler = new SequenceHandler(
            Enumerable.Range(0, 10).Select(i => ChatPage(i)).ToArray());
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var service = new GraphTeamsService(
            new GraphServiceClient(requestAdapter),
            Options.Create(new GraphAuthOptions()));

        var chats = await service.ListChatsAsync();

        chats.Should().HaveCount(25); // default top, taken from the newest page
        handler.Requests.Should().HaveCount(10); // budget spent: no 11th request
    }

    [Fact]
    public async Task ListChats_follows_an_empty_page_that_still_carries_a_continuation()
    {
        var handler = new SequenceHandler(
            Response("""
                {"value":[],"@odata.nextLink":"https://graph.microsoft.com/v1.0/me/chats?$skiptoken=next"}
                """),
            Response("""
                {"value":[{"id":"chat-1","chatType":"oneOnOne","lastUpdatedDateTime":"2026-11-25T09:00:00.000Z"}]}
                """));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var service = new GraphTeamsService(
            new GraphServiceClient(requestAdapter),
            Options.Create(new GraphAuthOptions()));

        var chats = await service.ListChatsAsync();

        chats.Should().ContainSingle(chat => chat.Id == "chat-1");
        handler.Requests.Should().HaveCount(2);
    }

    [Fact]
    public async Task ListChatMembers_maps_identity_and_follows_the_continuation()
    {
        var handler = new SequenceHandler(
            Response("""
                {"value":[{"@odata.type":"#microsoft.graph.aadUserConversationMember","id":"m-1","displayName":"Bob","email":"bob@contoso.com","userId":"u-bob","roles":[]}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/me/chats/chat-1/members?$skiptoken=next"}
                """),
            Response("""
                {"value":[{"@odata.type":"#microsoft.graph.aadUserConversationMember","id":"m-2","displayName":"Carol","email":"carol@contoso.com","userId":"u-carol","roles":["guest"]}]}
                """));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var service = new GraphTeamsService(
            new GraphServiceClient(requestAdapter),
            Options.Create(new GraphAuthOptions()));

        var members = await service.ListChatMembersAsync("chat-1");

        members.Select(m => m.DisplayName).Should().Equal("Bob", "Carol");
        members[0].Email.Should().Be("bob@contoso.com");
        members[0].UserId.Should().Be("u-bob");
        members[0].Roles.Should().BeEmpty();
        members[1].Roles.Should().Equal("guest");
        handler.Requests.Should().HaveCount(2);
        Uri.UnescapeDataString(handler.Requests[0].RequestUri!.AbsolutePath)
            .Should().Be("/v1.0/me/chats/chat-1/members");
        // The endpoint documents no OData query parameters.
        handler.Requests[0].RequestUri!.Query.Should().BeEmpty();
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
    public async Task GetChannelFilesFolder_uses_documented_path_without_select_and_maps_drive()
    {
        var handler = new SequenceHandler(
            Response("""
                {"id":"folder-1","name":"General","webUrl":"https://contoso.sharepoint.com/folder","parentReference":{"driveId":"drive-1","siteId":"site-1"}}
                """));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var service = new GraphTeamsService(
            new GraphServiceClient(requestAdapter),
            Options.Create(new GraphAuthOptions()));

        var folder = await service.GetChannelFilesFolderAsync("team-1", "channel-1");

        folder.TeamId.Should().Be("team-1");
        folder.ChannelId.Should().Be("channel-1");
        folder.DriveId.Should().Be("drive-1");
        folder.FolderId.Should().Be("folder-1");
        folder.SiteId.Should().Be("site-1");
        handler.Requests.Should().HaveCount(1);
        handler.Requests[0].RequestUri!.AbsolutePath.Should().Be(
            "/v1.0/teams/team-1/channels/channel-1/filesFolder");
        // The filesFolder endpoint documents no OData query parameters.
        handler.Requests[0].RequestUri!.Query.Should().BeEmpty();
    }

    [Fact]
    public async Task GetChannelFilesFolder_rejects_response_without_drive_id()
    {
        var handler = new SequenceHandler(
            Response("""{"id":"folder-1","name":"General"}"""));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var service = new GraphTeamsService(
            new GraphServiceClient(requestAdapter),
            Options.Create(new GraphAuthOptions()));

        var ex = await Assert.ThrowsAsync<GraphServiceException>(
            () => service.GetChannelFilesFolderAsync("team-1", "channel-1"));

        ex.Code.Should().Be("invalid-request");
    }

    private static HttpResponseMessage Response(string json) => new(HttpStatusCode.OK)
    {
        Content = new StringContent(json, Encoding.UTF8, "application/json")
    };

    /// <summary>A full $top=50 chat page with an optional @odata.nextLink.</summary>
    private static HttpResponseMessage ChatPage(int page, bool more = true)
    {
        var value = Enumerable.Range(0, 50)
            .Select(n => new
            {
                id = $"chat-{page}-{n}",
                chatType = "oneOnOne",
                lastUpdatedDateTime = $"2026-01-{page + 1:D2}T00:00:00.000Z"
            })
            .ToArray();
        var payload = new Dictionary<string, object?> { ["value"] = value };
        if (more)
        {
            payload["@odata.nextLink"] =
                $"https://graph.microsoft.com/v1.0/me/chats?$skiptoken=p{page + 1}";
        }

        return Response(JsonSerializer.Serialize(payload));
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