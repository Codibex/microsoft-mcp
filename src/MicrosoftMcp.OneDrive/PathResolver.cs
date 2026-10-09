namespace MicrosoftMcp.OneDrive;

/// <summary>Item references: "root", a Graph item id, or a /path/from/root.
/// Pure and unit-tested.</summary>
internal static class PathResolver
{
    internal static bool IsPath(string itemRef) =>
        !string.IsNullOrWhiteSpace(itemRef) && itemRef.TrimStart().StartsWith('/');

    internal static bool IsRoot(string itemRef) =>
        string.Equals(itemRef.Trim(), "root", StringComparison.OrdinalIgnoreCase);

    internal static string NormalizePath(string itemRef) => itemRef.Trim().TrimStart('/');

    internal static string RequireRef(string itemRef, string what = "itemRef")
    {
        if (string.IsNullOrWhiteSpace(itemRef))
        {
            throw MicrosoftMcp.Common.GraphServiceException.InvalidRequest(
                $"{what} must not be empty.",
                "use \"root\", an item id from onedrive_list_children/search, or a /path/from/root");
        }

        return itemRef.Trim();
    }

    /// <summary>Host paths for path-based up/download. Must be absolute so a
    /// tool call can only address explicit locations (no CWD surprises).</summary>
    internal static string RequireAbsolutePath(string localPath, string what = "localPath")
    {
        if (string.IsNullOrWhiteSpace(localPath) || !Path.IsPathFullyQualified(localPath.Trim()))
        {
            throw MicrosoftMcp.Common.GraphServiceException.InvalidRequest(
                $"{what} must be an absolute path.",
                "pass a full path, e.g. \"/home/user/datei.pdf\" or \"C:\\Daten\\datei.pdf\"");
        }

        return Path.GetFullPath(localPath.Trim());
    }
}
