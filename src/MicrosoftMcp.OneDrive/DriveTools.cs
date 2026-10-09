using System.ComponentModel;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using MicrosoftMcp.Common;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace MicrosoftMcp.OneDrive;

public static class DriveServiceRegistration
{
    public static IServiceCollection AddOneDrive(this IServiceCollection services)
    {
        services.AddSingleton<IGraphDriveService, GraphDriveService>();
        return services;
    }
}

[McpServerToolType]
public sealed class DriveTools(IGraphDriveService drive, ILogger<DriveTools> log)
{
    private async Task<CallToolResult> InvokeAsync<T>(
        string operation, Func<Task<T>> call, string resource = "drive")
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

    [McpServerTool, Description("Show the default OneDrive (id, name, quota total/used/remaining).")]
    public Task<CallToolResult> onedrive_get_drive(CancellationToken ct = default) =>
        InvokeAsync("onedrive_get_drive", () => drive.GetDriveAsync(ct));

    [McpServerTool, Description("List accessible drives (own OneDrive, shared libraries).")]
    public Task<CallToolResult> onedrive_list_drives(CancellationToken ct = default) =>
        InvokeAsync("onedrive_list_drives", () => drive.ListDrivesAsync(ct));

    [McpServerTool, Description("Get metadata of one item (file or folder).")]
    public Task<CallToolResult> onedrive_get_item(
        [Description("Item reference: \"root\", item id, or /path/from/root")] string itemRef,
        CancellationToken ct = default) =>
        InvokeAsync("onedrive_get_item", () => drive.GetItemAsync(itemRef, ct));

    [McpServerTool, Description("List children of a folder (default root). Folders first, then by name.")]
    public Task<CallToolResult> onedrive_list_children(
        [Description("Folder reference: \"root\", folder id, or /path")] string folderRef = "root",
        [Description("Max items 1-200")] int top = 50,
        CancellationToken ct = default) =>
        InvokeAsync("onedrive_list_children", () => drive.ListChildrenAsync(folderRef, top, ct));

    [McpServerTool, Description("Search files and folders by name/keyword across the drive.")]
    public Task<CallToolResult> onedrive_search_files(
        [Description("Filename or keyword, e.g. \"Rechnung\"")] string query,
        [Description("Max results 1-50")] int top = 25,
        CancellationToken ct = default) =>
        InvokeAsync("onedrive_search_files", () => drive.SearchAsync(query, top, ct));

    [McpServerTool, Description("Download a file. Without localPath returns text decoded (truncated at 20000 chars) or binary as base64; files above maxBytes are rejected. With localPath (absolute host path) streams the bytes to disk instead, so large files bypass the model context.")]
    public Task<CallToolResult> onedrive_download_file(
        [Description("Item reference: item id or /path/to/file")] string itemRef,
        [Description("Max bytes to download, up to 2097152")] int maxBytes = 786432,
        [Description("Absolute host path to save the file to (optional, for large files)")] string? localPath = null,
        [Description("Replace the local file if it exists")] bool overwrite = false,
        CancellationToken ct = default) =>
        InvokeAsync("onedrive_download_file", () => localPath is null
            ? drive.DownloadAsync(itemRef, maxBytes, ct)
            : drive.DownloadToFileAsync(itemRef, localPath, overwrite, ct));

    [McpServerTool, Description("Create a drive folder (existing names get a renamed copy, no error).")]
    public Task<CallToolResult> onedrive_create_folder(
        [Description("Plain folder name, e.g. \"Belege\"")] string name,
        [Description("Parent reference: \"root\", folder id, or /path")] string parentRef = "root",
        CancellationToken ct = default) =>
        InvokeAsync("onedrive_create_folder", () => drive.CreateFolderAsync(name, parentRef, ct));

    [McpServerTool, Description("Upload a file. Pass localPath (absolute host path, read from disk so bytes bypass the model context) or small inline content. Up to 4194304 bytes use simple upload, larger local files (up to 104857600 bytes) a resumable session. Existing names get a renamed copy, never overwritten.")]
    public Task<CallToolResult> onedrive_upload_file(
        [Description("Plain file name, e.g. \"notiz.txt\"")] string fileName,
        [Description("Parent reference: \"root\", folder id, or /path")] string? parentRef = null,
        [Description("Text content (alternative to localPath/contentBase64)")] string? contentText = null,
        [Description("Base64 content for binary files")] string? contentBase64 = null,
        [Description("Absolute host path to read bytes from (preferred for larger files)")] string? localPath = null,
        CancellationToken ct = default) =>
        InvokeAsync("onedrive_upload_file", () => drive.UploadAsync(fileName, parentRef, contentText, contentBase64, localPath, ct));

    [McpServerTool, Description("Move and/or rename an item. Reversible by moving back. No delete offered.")]
    public Task<CallToolResult> onedrive_move_item(
        [Description("Item id to move")] string itemId,
        [Description("Destination folder: id or /path (optional if renaming)")] string? newParentRef = null,
        [Description("New name (optional if moving)")] string? newName = null,
        CancellationToken ct = default) =>
        InvokeAsync("onedrive_move_item", () => drive.MoveAsync(itemId, newParentRef, newName, ct));
}
