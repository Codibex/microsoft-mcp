using AwesomeAssertions;
using Microsoft.Graph.Models;
using System.Text.Json;

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

        [Fact]
        public void MapInsightDetail_reads_notes_actions_and_mentions()
        {
                using var document = JsonDocument.Parse("""
                        {
                            "id": "i1",
                            "callId": "c1",
                            "contentCorrelationId": "corr1",
                            "createdDateTime": "2025-01-15T10:00:00Z",
                            "endDateTime": "2025-01-15T10:30:00Z",
                            "meetingNotes": [{
                                "title": "Decision",
                                "text": "Use Graph",
                                "subpoints": [{"title": "Next", "text": "Implement"}]
                            }],
                            "actionItems": [{"title": "Implement", "text": "Add tools", "ownerDisplayName": "Alice"}],
                            "viewpoint": {"mentionEvents": [{
                                "eventDateTime": "2025-01-15T10:10:00Z",
                                "transcriptUtterance": "Alice mentioned Bob.",
                                "speaker": {"user": {"displayName": "Alice"}}
                            }]}
                        }
                        """);

                var detail = TeamsMapper.MapInsightDetail(document.RootElement);

                detail.MeetingNotes.Should().ContainSingle(n => n.Subpoints.Single().Text == "Implement");
                detail.ActionItems.Should().ContainSingle(a => a.OwnerDisplayName == "Alice");
                detail.Mentions.Should().ContainSingle(m => m.Speaker == "Alice");
        }
}
