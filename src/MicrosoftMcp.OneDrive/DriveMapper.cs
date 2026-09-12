using Microsoft.Graph.Models;

namespace MicrosoftMcp.OneDrive;

/// <summary>Pure mapping helpers (Graph models to DTOs). Internal but unit-tested.</summary>
internal static class DriveMapper
{
    internal static DriveInfoDto MapDrive(Drive d) => new(
        d.Id ?? string.Empty,
        d.Name ?? string.Empty,
        d.DriveType,
        d.Quota?.Total ?? 0,
        d.Quota?.Used ?? 0,
        d.Quota?.Remaining ?? 0);

    internal static DriveItemSummary MapItem(DriveItem item) => new(
        item.Id ?? string.Empty,
        item.Name ?? string.Empty,
        item.Folder is not null,
        item.Size ?? 0,
        item.File?.MimeType,
        item.ParentReference?.Id,
        item.WebUrl,
        item.LastModifiedDateTime,
        item.Folder?.ChildCount ?? 0);

    internal static string? Truncate(string? s, int max) =>
        s is null || s.Length <= max ? s : s[..max] + "…[truncated]";

    internal static bool IsTextContent(string? mimeType)
    {
        if (string.IsNullOrWhiteSpace(mimeType))
        {
            return false;
        }

        string type = mimeType.Split(';')[0].Trim().ToLowerInvariant();
        return type.StartsWith("text/", StringComparison.Ordinal)
            || type is "application/json" or "application/xml"
                or "application/javascript" or "application/csv"
            || type.EndsWith("+json", StringComparison.Ordinal)
            || type.EndsWith("+xml", StringComparison.Ordinal);
    }
}
