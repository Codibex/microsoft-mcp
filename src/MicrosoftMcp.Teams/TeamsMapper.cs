using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Graph.Models;

namespace MicrosoftMcp.Teams;

/// <summary>Pure mapping helpers (Graph models to DTOs). Internal but unit-tested.</summary>
internal static partial class TeamsMapper
{
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
        c.LastUpdatedDateTime,
        [.. (c.Members ?? []).Select(MapChatMember)]);

    /// <summary>email/userId live on the aad subtype; displayName and roles
    /// live on the base conversationMember type (present for guest/external
    /// members as well).</summary>
    internal static ChatMemberInfo MapChatMember(ConversationMember m) =>
        m is AadUserConversationMember aad
            ? new ChatMemberInfo(m.DisplayName, aad.Email, aad.UserId, m.Roles)
            : new ChatMemberInfo(m.DisplayName, null, null, m.Roles);

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

    internal static MessageSearchResult MapSearchResult(JsonElement root)
    {
        if (!root.TryGetProperty("value", out var values) || values.ValueKind != JsonValueKind.Array)
        {
            return new MessageSearchResult([], 0, false);
        }

        var hits = new List<MessageSearchHit>();
        int total = 0;
        bool moreResultsAvailable = false;

        foreach (JsonElement response in values.EnumerateArray())
        {
            if (!response.TryGetProperty("hitsContainers", out var containers)
                || containers.ValueKind != JsonValueKind.Array)
            {
                continue;
            }

            foreach (JsonElement container in containers.EnumerateArray())
            {
                if (container.TryGetProperty("total", out var totalEl) && totalEl.TryGetInt32(out int t))
                {
                    total += t;
                }

                if (container.TryGetProperty("moreResultsAvailable", out var moreEl)
                    && moreEl.ValueKind == JsonValueKind.True)
                {
                    moreResultsAvailable = true;
                }

                if (!container.TryGetProperty("hits", out var hitsEl) || hitsEl.ValueKind != JsonValueKind.Array)
                {
                    continue;
                }

                foreach (JsonElement hit in hitsEl.EnumerateArray())
                {
                    hits.Add(MapSearchHit(hit));
                }
            }
        }

        return new MessageSearchResult(hits, total, moreResultsAvailable);
    }

    internal static MessageSearchHit MapSearchHit(JsonElement hit)
    {
        bool hasResource = hit.TryGetProperty("resource", out var resource)
            && resource.ValueKind == JsonValueKind.Object;

        string messageId = (hasResource ? StringProperty(resource, "id") : null)
            ?? StringProperty(hit, "hitId")
            ?? string.Empty;

        string? chatId = hasResource ? NonEmpty(StringProperty(resource, "chatId")) : null;
        string? teamId = null;
        string? channelId = null;
        if (hasResource
            && resource.TryGetProperty("channelIdentity", out var channelIdentity)
            && channelIdentity.ValueKind == JsonValueKind.Object)
        {
            teamId = NonEmpty(StringProperty(channelIdentity, "teamId"));
            channelId = NonEmpty(StringProperty(channelIdentity, "channelId"));
        }

        (string? from, string? fromAddress) = hasResource ? ReadSearchSender(resource) : (null, null);
        DateTimeOffset? created = hasResource
            ? DateProperty(resource, "createdDateTime") ?? DateProperty(resource, "lastModifiedDateTime")
            : null;
        string? summary = Truncate(StripHtml(StringProperty(hit, "summary")), 500);
        string? webUrl = hasResource
            ? NonEmpty(StringProperty(resource, "webUrl")) ?? NonEmpty(StringProperty(resource, "webLink"))
            : null;

        return new MessageSearchHit(
            messageId,
            chatId,
            teamId,
            channelId,
            from,
            fromAddress,
            created,
            summary,
            webUrl);
    }

    private static (string? Name, string? Address) ReadSearchSender(JsonElement resource)
    {
        if (!resource.TryGetProperty("from", out var from) || from.ValueKind != JsonValueKind.Object)
        {
            return (null, null);
        }

        // Microsoft Search returns resource.from.emailAddress.{name, address}
        // instead of the standard chatMessageFromIdentitySet.user shape.
        if (from.TryGetProperty("emailAddress", out var emailAddress)
            && emailAddress.ValueKind == JsonValueKind.Object)
        {
            return (
                NonEmpty(StringProperty(emailAddress, "name")),
                NonEmpty(StringProperty(emailAddress, "address")));
        }

        foreach (string identity in new[] { "user", "application", "device" })
        {
            if (from.TryGetProperty(identity, out var value) && value.ValueKind == JsonValueKind.Object)
            {
                string? displayName = NonEmpty(StringProperty(value, "displayName"));
                if (displayName is not null)
                {
                    return (displayName, null);
                }
            }
        }

        return (null, null);
    }

    private static string? NonEmpty(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;

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
