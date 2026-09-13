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

    internal static string? SenderName(ChatMessage m) =>
        m.From?.User?.DisplayName
        ?? m.From?.Application?.DisplayName
        ?? m.From?.Device?.DisplayName;

    internal static string? StripHtml(string? html) =>
        html is null ? null : System.Net.WebUtility.HtmlDecode(TagsRegex().Replace(html, " ")).Trim();

    internal static string? Truncate(string? s, int max) =>
        s is null || s.Length <= max ? s : s[..max] + "…[truncated]";

    [GeneratedRegex("<[^>]+>")]
    private static partial Regex TagsRegex();
}
