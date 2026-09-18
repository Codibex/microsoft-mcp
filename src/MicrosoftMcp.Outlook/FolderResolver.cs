using MicrosoftMcp.Common;

namespace MicrosoftMcp.Outlook;

/// <summary>Pure folder-destination resolution. Unit-tested without Graph.</summary>
internal static class FolderResolver
{
    private static readonly HashSet<string> WellKnown = new(StringComparer.OrdinalIgnoreCase)
    {
        "inbox", "archive", "deleteditems", "drafts", "sentitems", "junkemail", "outbox"
    };

    internal static bool IsWellKnown(string destination) =>
        WellKnown.Contains(destination.Trim());

    internal static string Resolve(string destination, IReadOnlyList<FolderInfo> folders)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(destination);
        string trimmed = destination.Trim();

        // Well-known names pass through: Graph resolves them by name.
        if (IsWellKnown(trimmed))
        {
            return trimmed.ToLowerInvariant();
        }

        var byId = folders.FirstOrDefault(f =>
            string.Equals(f.Id, trimmed, StringComparison.OrdinalIgnoreCase));
        if (byId is not null)
        {
            return byId.Id;
        }

        var byName = folders.FirstOrDefault(f =>
            string.Equals(f.DisplayName, trimmed, StringComparison.OrdinalIgnoreCase));
        if (byName is not null)
        {
            return byName.Id;
        }

        throw GraphServiceException.FolderNotFound(destination);
    }
}
