using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Graph.Models;

namespace MicrosoftMcp.Teams;

/// <summary>Pure mapping helpers (Graph models to DTOs). Internal but unit-tested.</summary>
internal static partial class TeamsMapper
{
    private static readonly string[] MessageSelect =
        ["id", "messageType", "body", "from", "createdDateTime", "attachments",
         "mentions", "reactions", "replyToId", "webUrl"];

    internal static string[] Select => MessageSelect;

    internal static TeamInfo MapTeam(Team t) => new(
        t.Id ?? string.Empty,
        t.DisplayName ?? string.Empty,
        t.Description,
        t.Visibility?.ToString(),
        t.IsArchived ?? false);

    internal static ChannelInfo MapChannel(Channel c) => new(
        c.Id ?? string.Empty,
        c.DisplayName ?? string.Empty,
        c.Description,
        c.MembershipType?.ToString());

    internal static ChatInfo MapChat(Chat c) => new(
        c.Id ?? string.Empty,
        c.Topic,
        c.ChatType?.ToString(),
        c.LastUpdatedDateTime);

    internal static MessageSummary MapSummary(ChatMessage m) => new(
        m.Id ?? string.Empty,
        m.MessageType?.ToString(),
        SenderName(m),
        m.CreatedDateTime,
        Truncate(StripHtml(m.Body?.Content), 500),
        m.Body?.ContentType?.ToString(),
        m.Reactions?.Count ?? 0,
        m.Attachments?.Count ?? 0,
        m.ReplyToId,
        m.WebUrl);

    internal static MessageDetail MapDetail(ChatMessage m) => new(
        m.Id ?? string.Empty,
        m.MessageType?.ToString(),
        SenderName(m),
        m.CreatedDateTime,
        Truncate(m.Body?.Content, 8000),
        m.Body?.ContentType?.ToString(),
        [.. (m.Reactions ?? []).Select(r => new ReactionDto(
            r.ReactionType,
            r.DisplayName,
            r.CreatedDateTime))],
        [.. (m.Mentions ?? []).Select(x => new MentionDto(
            x.MentionText,
            x.Mentioned?.User?.DisplayName
                ?? x.Mentioned?.Application?.DisplayName))],
        m.Attachments?.Count ?? 0,
        m.ReplyToId,
        m.WebUrl);

    internal static MeetingTranscriptInfo MapTranscript(CallTranscript transcript) => new(
        transcript.Id ?? string.Empty,
        transcript.MeetingId,
        transcript.CallId,
        transcript.CreatedDateTime,
        transcript.EndDateTime,
        transcript.ContentCorrelationId);

    internal static MeetingTranscriptDetail MapTranscriptDetail(CallTranscript transcript, string content) => new(
        transcript.Id ?? string.Empty,
        transcript.MeetingId,
        transcript.CallId,
        transcript.CreatedDateTime,
        transcript.EndDateTime,
        transcript.ContentCorrelationId,
        Truncate(content, 100_000) ?? string.Empty,
        "text/vtt");

    internal static MeetingInsightInfo MapInsightSummary(JsonElement insight) => new(
        StringProperty(insight, "id") ?? string.Empty,
        StringProperty(insight, "callId"),
        StringProperty(insight, "contentCorrelationId"),
        DateProperty(insight, "createdDateTime"),
        DateProperty(insight, "endDateTime"));

    internal static MeetingInsightDetail MapInsightDetail(JsonElement insight) => new(
        StringProperty(insight, "id") ?? string.Empty,
        StringProperty(insight, "callId"),
        StringProperty(insight, "contentCorrelationId"),
        DateProperty(insight, "createdDateTime"),
        DateProperty(insight, "endDateTime"),
        ReadNotes(insight),
        ReadActionItems(insight),
        ReadMentions(insight));

    internal static string? SenderName(ChatMessage m) =>
        m.From?.User?.DisplayName
        ?? m.From?.Application?.DisplayName
        ?? m.From?.Device?.DisplayName;

    internal static string? StripHtml(string? html) =>
        html is null ? null : System.Net.WebUtility.HtmlDecode(TagsRegex().Replace(html, " ")).Trim();

    internal static string? Truncate(string? s, int max) =>
        s is null || s.Length <= max ? s : s[..max] + "…[truncated]";

    private static string? StringProperty(JsonElement element, string name) =>
        element.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.String
            ? property.GetString()
            : null;

    private static DateTimeOffset? DateProperty(JsonElement element, string name)
    {
        string? value = StringProperty(element, name);
        return DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var result)
            ? result
            : null;
    }

    private static IReadOnlyList<MeetingNoteInfo> ReadNotes(JsonElement insight) =>
        ArrayProperty(insight, "meetingNotes")
            .Select(note => new MeetingNoteInfo(
                StringProperty(note, "title"),
                StringProperty(note, "text"),
                ReadSubpoints(note)))
            .ToArray();

    private static IReadOnlyList<MeetingNoteSubpointInfo> ReadSubpoints(JsonElement note) =>
        ArrayProperty(note, "subpoints")
            .Select(subpoint => new MeetingNoteSubpointInfo(
                StringProperty(subpoint, "title"),
                StringProperty(subpoint, "text")))
            .ToArray();

    private static IReadOnlyList<MeetingActionItemInfo> ReadActionItems(JsonElement insight) =>
        ArrayProperty(insight, "actionItems")
            .Select(item => new MeetingActionItemInfo(
                StringProperty(item, "title"),
                StringProperty(item, "text"),
                StringProperty(item, "ownerDisplayName")))
            .ToArray();

    private static IReadOnlyList<MeetingMentionInfo> ReadMentions(JsonElement insight) =>
        ArrayProperty(insight, "viewpoint")
            .SelectMany(viewpoint => ArrayProperty(viewpoint, "mentionEvents"))
            .Select(mention => new MeetingMentionInfo(
                DateProperty(mention, "eventDateTime"),
                StringProperty(mention, "transcriptUtterance"),
                ReadSpeaker(mention)))
            .ToArray();

    private static IEnumerable<JsonElement> ArrayProperty(JsonElement element, string name)
    {
        if (!element.TryGetProperty(name, out var property))
        {
            return [];
        }

        return property.ValueKind == JsonValueKind.Array
            ? property.EnumerateArray()
            : [property];
    }

    private static string? ReadSpeaker(JsonElement mention)
    {
        if (!mention.TryGetProperty("speaker", out var speaker))
        {
            return null;
        }

        foreach (string identity in new[] { "user", "application", "device" })
        {
            if (speaker.TryGetProperty(identity, out var value))
            {
                string? displayName = StringProperty(value, "displayName");
                if (!string.IsNullOrWhiteSpace(displayName))
                {
                    return displayName;
                }
            }
        }

        return null;
    }

    [GeneratedRegex("<[^>]+>")]
    private static partial Regex TagsRegex();
}
