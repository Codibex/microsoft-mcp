namespace MicrosoftMcp.SharePoint;

/// <summary>Item references: "root", a Graph item id, or a /path/from/root.
/// Drive ids are always explicit: unlike the fixed-drive onedrive_* tools,
/// every sharepoint_* file tool takes the driveId (from
/// sharepoint_list_site_drives or teams_get_channel_files_folder).
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
                "use \"root\", an item id from sharepoint_list_children/search, or a /path/from/root");
        }

        return itemRef.Trim();
    }

    internal static string RequireDriveId(string driveId)
    {
        if (string.IsNullOrWhiteSpace(driveId))
        {
            throw MicrosoftMcp.Common.GraphServiceException.InvalidRequest(
                "driveId must not be empty.",
                "call sharepoint_list_site_drives with the site id, or teams_get_channel_files_folder for a channel, to get valid drive ids");
        }

        return driveId.Trim();
    }

    internal static string RequireSiteId(string siteId)
    {
        if (string.IsNullOrWhiteSpace(siteId))
        {
            throw MicrosoftMcp.Common.GraphServiceException.InvalidRequest(
                "siteId must not be empty.",
                "call sharepoint_search_sites to get valid site ids (or \"root\" for the tenant root site)");
        }

        return siteId.Trim();
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
