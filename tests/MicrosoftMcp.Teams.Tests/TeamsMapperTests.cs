using AwesomeAssertions;
using Microsoft.Graph.Models;

namespace MicrosoftMcp.Teams.Tests;

public sealed class TeamsMapperTests
{
    [Fact]
    public void MapTeam_is_null_safe()
    {
        var info = TeamsMapper.MapTeam(new Team());

        info.Id.Should().BeEmpty();
        info.IsArchived.Should().BeFalse();
    }

    [Fact]
    public void MapSummary_strips_html_for_preview()
    {
        var summary = TeamsMapper.MapSummary(new ChatMessage
        {
            Id = "m1",
            Body = new ItemBody { Content = "<p>Hallo <b>Welt</b></p>", ContentType = BodyType.Html },
            From = new ChatMessageFromIdentitySet
            {
                User = new TeamworkUserIdentity { DisplayName = "Alice" }
            }
        });

        summary.From.Should().Be("Alice");
        summary.Preview.Should().Contain("Hallo");
        summary.Preview.Should().NotContain("<p>");
    }

    [Fact]
    public void SenderName_falls_back_through_identities()
    {
        TeamsMapper.SenderName(new ChatMessage()).Should().BeNull();
        TeamsMapper.SenderName(new ChatMessage
        {
            From = new ChatMessageFromIdentitySet
            {
                Application = new TeamworkApplicationIdentity { DisplayName = "Bot" }
            }
        }).Should().Be("Bot");
    }

    [Fact]
    public void MapDetail_maps_reactions_and_mentions()
    {
        var detail = TeamsMapper.MapDetail(new ChatMessage
        {
            Id = "m1",
            Reactions =
            [
                new ChatMessageReaction { ReactionType = "like", DisplayName = "Bob" }
            ],
            Mentions =
            [
                new ChatMessageMention
                {
                    MentionText = "@Alice",
                    Mentioned = new ChatMessageMentionedIdentitySet
                    {
                        User = new TeamworkUserIdentity { DisplayName = "Alice" }
                    }
                }
            ]
        });

        detail.Reactions.Should().ContainSingle(r => r.Type == "like");
        detail.Mentions.Should().ContainSingle(x => x.Mentioned == "Alice");
    }
}
