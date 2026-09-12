using Microsoft.Graph.Models;

namespace MicrosoftMcp.Outlook;

/// <summary>Pure mapping helpers (Graph models to DTOs). Internal but unit-tested.</summary>
internal static class EmailMapper
{
    internal static EmailSummary MapSummary(Message m) => new(
        m.Id ?? string.Empty,
        m.Subject ?? string.Empty,
        m.From?.EmailAddress is { } e ? new EmailAddressDto(e.Name ?? string.Empty, e.Address ?? string.Empty) : null,
        [.. (m.ToRecipients ?? []).Select(r => new EmailAddressDto(
            r.EmailAddress?.Name ?? string.Empty, r.EmailAddress?.Address ?? string.Empty))],
        m.ReceivedDateTime,
        m.IsRead ?? false,
        m.HasAttachments ?? false,
        m.Categories ?? [],
        m.Importance?.ToString(),
        Truncate(m.BodyPreview, 300));

    internal static EmailDetail MapDetail(Message m) => new(
        m.Id ?? string.Empty,
        m.Subject ?? string.Empty,
        m.From?.EmailAddress is { } e ? new EmailAddressDto(e.Name ?? string.Empty, e.Address ?? string.Empty) : null,
        [.. (m.ToRecipients ?? []).Select(r => new EmailAddressDto(
            r.EmailAddress?.Name ?? string.Empty, r.EmailAddress?.Address ?? string.Empty))],
        m.ReceivedDateTime,
        m.IsRead ?? false,
        m.Categories ?? [],
        Truncate(m.BodyPreview, 500),
        Truncate(m.Body?.Content, 8000),
        m.WebLink);

    internal static AttachmentInfo MapAttachment(Attachment a) => new(
        a.Id ?? string.Empty,
        a.Name ?? string.Empty,
        a.ContentType,
        a.Size ?? 0,
        a.IsInline ?? false);

    internal static CategoryInfo MapCategory(OutlookCategory c) => new(
        c.Id ?? string.Empty,
        c.DisplayName ?? string.Empty,
        c.Color?.ToString());

    internal static string? Truncate(string? s, int max) =>
        s is null || s.Length <= max ? s : s[..max] + "…[truncated]";

    internal static List<string> MergeCategories(
        IReadOnlyList<string> current, string[] add, string[] remove) =>
        [.. current
            .Where(c => !remove.Contains(c, StringComparer.OrdinalIgnoreCase))
            .Concat(add)
            .Distinct(StringComparer.OrdinalIgnoreCase)];
}
