using MicrosoftMcp.Common;

namespace MicrosoftMcp.Outlook;

/// <summary>Pure folder-destination resolution. Unit-tested without Graph.</summary>
internal static class FolderResolver
{
    private static readonly HashSet<string> WellKnown = new(StringComparer.OrdinalIgnoreCase)
    {
        "archive", "clutter", "conflicts", "conversationhistory", "deleteditems", "drafts",
        "inbox", "junkemail", "localfailures", "msgfolderroot", "outbox",
        "recoverableitemsdeletions", "scheduled", "searchfolders", "sentitems",
        "serverfailures", "syncissues"
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

        var byPath = folders.Where(f =>
            !string.IsNullOrWhiteSpace(f.Path)
            && string.Equals(f.Path, trimmed, StringComparison.OrdinalIgnoreCase)).ToList();
        if (byPath.Count == 1)
        {
            return byPath[0].Id;
        }

        if (byPath.Count > 1)
        {
            throw Ambiguous(trimmed);
        }

        var byName = folders.Where(f =>
            string.Equals(f.DisplayName, trimmed, StringComparison.OrdinalIgnoreCase));
        if (byName.Count() == 1)
        {
            return byName.Single().Id;
        }

        if (byName.Count() > 1)
        {
            throw Ambiguous(trimmed);
        }

        throw GraphServiceException.FolderNotFound(destination);
    }

    private static GraphServiceException Ambiguous(string destination) =>
        GraphServiceException.InvalidRequest(
            $"Folder destination '{destination}' is ambiguous.",
            "use the full folder path from outlook_list_folders or the folder id");
}
