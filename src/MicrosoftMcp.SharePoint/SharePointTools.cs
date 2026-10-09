using System.ComponentModel;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using MicrosoftMcp.Common;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace MicrosoftMcp.SharePoint;

public static class SharePointServiceRegistration
{
    public static IServiceCollection AddSharePoint(this IServiceCollection services)
    {
        services.AddSingleton<IGraphSharePointService, GraphSharePointService>();
        return services;
    }
}

/// <summary>SharePoint sites + file access on explicit drives. Clean split
/// from the fixed-drive onedrive_* tools: every file tool here takes the
/// driveId (from sharepoint_list_site_drives or
/// teams_get_channel_files_folder). No delete is offered.</summary>
[McpServerToolType]
public sealed class SharePointTools(IGraphSharePointService sites, ILogger<SharePointTools> log)
{
    private async Task<CallToolResult> InvokeAsync<T>(
        string operation, Func<Task<T>> call, string resource = "sharepoint-drive")
    {
        try
        {
            return ToolResult.Ok(await call().ConfigureAwait(false));
        }
        catch (GraphServiceException ex)
        {
            log.LogWarning(ex, "Tool {Operation} failed with {Code}", operation, ex.Code);
            return ToolResult.Fail(ex);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            var mapped = GraphErrorMapper.ToGraphServiceException(ex, operation, resource);
            log.LogError(ex, "Tool {Operation} failed unexpectedly ({Code})", operation, mapped.Code);
            return ToolResult.Fail(mapped);
        }
    }

    [McpServerTool, Description("Search SharePoint sites by name/keyword. Needed to pick a site id. Read-only.")]
    public Task<CallToolResult> sharepoint_search_sites(
        [Description("Site or team name, e.g. \"Engineering\"")] string query,
        [Description("Max results 1-50")] int top = 25,
        CancellationToken ct = default) =>
        InvokeAsync("sharepoint_search_sites", () => sites.SearchSitesAsync(query, top, ct), "sharepoint-site");

    [McpServerTool, Description("Get metadata of one site (\"root\" = tenant root site). Read-only.")]
    public Task<CallToolResult> sharepoint_get_site(
        [Description("Site id from sharepoint_search_sites, or \"root\"")] string siteId,
        CancellationToken ct = default) =>
        InvokeAsync("sharepoint_get_site", () => sites.GetSiteAsync(siteId, ct), "sharepoint-site");

    [McpServerTool, Description("List document libraries (drives) of a site. Needed to pick a driveId for the file tools. Read-only.")]
    public Task<CallToolResult> sharepoint_list_site_drives(
        [Description("Site id from sharepoint_search_sites or sharepoint_get_site")] string siteId,
        [Description("Max drives 1-100")] int top = 50,
        CancellationToken ct = default) =>
        InvokeAsync("sharepoint_list_site_drives", () => sites.ListSiteDrivesAsync(siteId, top, ct), "sharepoint-site");

    [McpServerTool, Description("Get metadata of one item (file or folder) on an explicit drive.")]
    public Task<CallToolResult> sharepoint_get_item(
        [Description("Drive id from sharepoint_list_site_drives or teams_get_channel_files_folder")] string driveId,
        [Description("Item reference: \"root\", item id, or /path/from/root")] string itemRef,
        CancellationToken ct = default) =>
        InvokeAsync("sharepoint_get_item", () => sites.GetItemAsync(driveId, itemRef, ct));

    [McpServerTool, Description("List children of a folder on an explicit drive (default root). Folders first, then by name.")]
    public Task<CallToolResult> sharepoint_list_children(
        [Description("Drive id from sharepoint_list_site_drives or teams_get_channel_files_folder")] string driveId,
        [Description("Folder reference: \"root\", folder id, or /path")] string folderRef = "root",
        [Description("Max items 1-200")] int top = 50,
        CancellationToken ct = default) =>
        InvokeAsync("sharepoint_list_children", () => sites.ListChildrenAsync(driveId, folderRef, top, ct));

    [McpServerTool, Description("Search files and folders by name/keyword on an explicit drive.")]
    public Task<CallToolResult> sharepoint_search_files(
        [Description("Drive id from sharepoint_list_site_drives or teams_get_channel_files_folder")] string driveId,
        [Description("Filename or keyword, e.g. \"Rechnung\"")] string query,
        [Description("Max results 1-50")] int top = 25,
        CancellationToken ct = default) =>
        InvokeAsync("sharepoint_search_files", () => sites.SearchAsync(driveId, query, top, ct));

    [McpServerTool, Description("Download a file from an explicit drive. Without localPath returns text decoded (truncated at 20000 chars) or binary as base64; files above maxBytes are rejected. With localPath (absolute host path) streams the bytes to disk instead, so large files bypass the model context.")]
    public Task<CallToolResult> sharepoint_download_file(
        [Description("Drive id from sharepoint_list_site_drives or teams_get_channel_files_folder")] string driveId,
        [Description("Item reference: item id or /path/to/file")] string itemRef,
        [Description("Max bytes to download, up to 2097152")] int maxBytes = 786432,
        [Description("Absolute host path to save the file to (optional, for large files)")] string? localPath = null,
        [Description("Replace the local file if it exists")] bool overwrite = false,
        CancellationToken ct = default) =>
        InvokeAsync("sharepoint_download_file", () => localPath is null
            ? sites.DownloadAsync(driveId, itemRef, maxBytes, ct)
            : sites.DownloadToFileAsync(driveId, itemRef, localPath, overwrite, ct));

    [McpServerTool, Description("Create a folder on an explicit drive (existing names get a renamed copy, no error).")]
    public Task<CallToolResult> sharepoint_create_folder(
        [Description("Drive id from sharepoint_list_site_drives or teams_get_channel_files_folder")] string driveId,
        [Description("Plain folder name, e.g. \"Belege\"")] string name,
        [Description("Parent reference: \"root\", folder id, or /path")] string parentRef = "root",
        CancellationToken ct = default) =>
        InvokeAsync("sharepoint_create_folder", () => sites.CreateFolderAsync(driveId, name, parentRef, ct));

    [McpServerTool, Description("Upload a file to an explicit drive. Pass localPath (absolute host path, read from disk so bytes bypass the model context) or small inline content. Up to 4194304 bytes use simple upload, larger local files (up to 104857600 bytes) a resumable session. Existing names get a renamed copy, never overwritten.")]
    public Task<CallToolResult> sharepoint_upload_file(
        [Description("Drive id from sharepoint_list_site_drives or teams_get_channel_files_folder")] string driveId,
        [Description("Plain file name, e.g. \"notiz.txt\"")] string fileName,
        [Description("Parent reference: \"root\", folder id, or /path")] string? parentRef = null,
        [Description("Text content (alternative to localPath/contentBase64)")] string? contentText = null,
        [Description("Base64 content for binary files")] string? contentBase64 = null,
        [Description("Absolute host path to read bytes from (preferred for larger files)")] string? localPath = null,
        CancellationToken ct = default) =>
        InvokeAsync("sharepoint_upload_file", () => sites.UploadAsync(driveId, fileName, parentRef, contentText, contentBase64, localPath, ct));

    [McpServerTool, Description("Move and/or rename an item on an explicit drive. Reversible by moving back. No delete offered.")]
    public Task<CallToolResult> sharepoint_move_item(
        [Description("Drive id from sharepoint_list_site_drives or teams_get_channel_files_folder")] string driveId,
        [Description("Item id to move")] string itemId,
        [Description("Destination folder: id or /path (optional if renaming)")] string? newParentRef = null,
        [Description("New name (optional if moving)")] string? newName = null,
        CancellationToken ct = default) =>
        InvokeAsync("sharepoint_move_item", () => sites.MoveAsync(driveId, itemId, newParentRef, newName, ct));
}
