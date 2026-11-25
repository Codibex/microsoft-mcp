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
    public async Task ListChatMembers_enforces_the_500_member_cap_within_a_page_and_stops()
    {
        // Two 300-member pages both advertising a continuation: the second page
        // must be truncated to the remaining 200 members and no 3rd page fetched.
        var handler = new SequenceHandler(
            MemberPage(0, count: 300),
            MemberPage(1, count: 300));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var service = new GraphTeamsService(
            new GraphServiceClient(requestAdapter),
            Options.Create(new GraphAuthOptions()));

        var members = await service.ListChatMembersAsync("chat-1");

        members.Should().HaveCount(500);
        members[0].DisplayName.Should().Be("Member 0-0");
        members[^1].DisplayName.Should().Be("Member 1-199");
        handler.Requests.Should().HaveCount(2);
    }

    [Fact]
    public async Task ListChatMembers_stops_at_the_page_cap_on_empty_continuations()
    {
        // Repeating empty pages with a non-empty @odata.nextLink must consume
        // the page budget instead of looping forever.
        var handler = new SequenceHandler(
            Enumerable.Range(0, 20).Select(i => MemberPage(i, count: 0)).ToArray());
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var service = new GraphTeamsService(
            new GraphServiceClient(requestAdapter),
            Options.Create(new GraphAuthOptions()));

        var members = await service.ListChatMembersAsync("chat-1");

        members.Should().BeEmpty();
        handler.Requests.Should().HaveCount(20);
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
    [Fact]
    public async Task SearchMessages_posts_chatMessage_query_and_maps_chat_and_channel_hits()
    {
        var handler = new SequenceHandler(
            Response("""
                {
                  "value": [
                    {
                      "hitsContainers": [
                        {
                          "hits": [
                            {
                              "hitId": "hit-chat",
                              "rank": 1,
                              "summary": "Let's discuss <c0>budget</c0> &amp; timeline",
                              "resource": {
                                "@odata.type": "microsoft.graph.chatMessage",
                                "id": "1657782060227",
                                "createdDateTime": "2026-11-20T09:00:00Z",
                                "chatId": "19:chat-abc@thread.v2",
                                "channelIdentity": {},
                                "from": {
                                  "emailAddress": {
                                    "name": "Alice Example",
                                    "address": "alice@contoso.com"
                                  }
                                },
                                "webLink": "https://teams.microsoft.com/l/message/chat-abc/1657782060227"
                              }
                            },
                            {
                              "hitId": "hit-channel",
                              "rank": 2,
                              "summary": "Channel <c0>budget</c0> update",
                              "resource": {
                                "@odata.type": "microsoft.graph.chatMessage",
                                "id": "1657782099999",
                                "createdDateTime": "2026-11-21T10:30:00Z",
                                "channelIdentity": {
                                  "teamId": "team-1",
                                  "channelId": "channel-1"
                                },
                                "from": {
                                  "user": {
                                    "displayName": "Bob"
                                  }
                                }
                              }
                            }
                          ],
                          "total": 42,
                          "moreResultsAvailable": true
                        }
                      ]
                    }
                  ]
                }
                """));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var service = new GraphTeamsService(
            new GraphServiceClient(requestAdapter),
            Options.Create(new GraphAuthOptions()));

        var result = await service.SearchMessagesAsync("  budget  ", top: 100);

        handler.Requests.Should().ContainSingle();
        handler.Requests[0].Method.Should().Be(HttpMethod.Post);
        handler.Requests[0].RequestUri!.AbsolutePath.Should().Be("/v1.0/search/query");

        using var sent = JsonDocument.Parse(handler.RequestBodies[0]);
        var req = sent.RootElement.GetProperty("requests")[0];
        req.GetProperty("entityTypes")[0].GetString().Should().Be("chatMessage");
        req.GetProperty("query").GetProperty("queryString").GetString().Should().Be("budget");
        req.GetProperty("from").GetInt32().Should().Be(0);
        req.GetProperty("size").GetInt32().Should().Be(50);

        result.Total.Should().Be(42);
        result.MoreResultsAvailable.Should().BeTrue();
        result.Hits.Should().HaveCount(2);

        var chatHit = result.Hits[0];
        chatHit.MessageId.Should().Be("1657782060227");
        chatHit.ChatId.Should().Be("19:chat-abc@thread.v2");
        chatHit.TeamId.Should().BeNull();
        chatHit.ChannelId.Should().BeNull();
        chatHit.From.Should().Be("Alice Example");
        chatHit.FromAddress.Should().Be("alice@contoso.com");
        chatHit.Summary.Should().Be("Let's discuss  budget  & timeline");
        chatHit.WebUrl.Should().Be("https://teams.microsoft.com/l/message/chat-abc/1657782060227");

        var channelHit = result.Hits[1];
        channelHit.MessageId.Should().Be("1657782099999");
        channelHit.ChatId.Should().BeNull();
        channelHit.TeamId.Should().Be("team-1");
        channelHit.ChannelId.Should().Be("channel-1");
        channelHit.From.Should().Be("Bob");
        channelHit.FromAddress.Should().BeNull();
    }

    [Fact]
    public async Task SearchMessages_handles_empty_hits_and_rejects_blank_query()
    {
        var handler = new SequenceHandler(
            Response("""{"value":[{"hitsContainers":[{"total":0,"moreResultsAvailable":false}]}]}"""));
        using var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var service = new GraphTeamsService(
            new GraphServiceClient(requestAdapter),
            Options.Create(new GraphAuthOptions()));

        var empty = await service.SearchMessagesAsync("nothing-matches");
        empty.Hits.Should().BeEmpty();
        empty.Total.Should().Be(0);
        empty.MoreResultsAvailable.Should().BeFalse();

        var act = () => service.SearchMessagesAsync("   ");
        var ex = await Assert.ThrowsAsync<GraphServiceException>(act);
        ex.Code.Should().Be("invalid-request");
    }

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

    /// <summary>A chat-members page with <paramref name="count"/> items and an optional @odata.nextLink.</summary>
    private static HttpResponseMessage MemberPage(int page, int count, bool more = true)
    {
        var value = Enumerable.Range(0, count)
            .Select(n => new Dictionary<string, object?>
            {
                ["@odata.type"] = "#microsoft.graph.aadUserConversationMember",
                ["id"] = $"m-{page}-{n}",
                ["displayName"] = $"Member {page}-{n}",
                ["email"] = $"member-{page}-{n}@contoso.com",
                ["userId"] = $"u-{page}-{n}",
                ["roles"] = Array.Empty<string>()
            })
            .ToArray();
        var payload = new Dictionary<string, object?> { ["value"] = value };
        if (more)
        {
            payload["@odata.nextLink"] =
                $"https://graph.microsoft.com/v1.0/me/chats/chat-1/members?$skiptoken=p{page + 1}";
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