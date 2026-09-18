using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using MicrosoftMcp.Common;

namespace MicrosoftMcp.OneDrive;

/// <summary>OneDrive access via the default drive (/me/drive). Read, create
/// and move only – no delete. Items are addressed by "root", id or /path.</summary>
public sealed class GraphDriveService : IGraphDriveService
{
    private static readonly string[] ItemSelect =
        ["id", "name", "size", "folder", "file", "parentReference", "webUrl", "lastModifiedDateTime"];

    // Simple upload limit (Graph simple upload caps at 4 MiB; larger files
    // need an upload session, which is out of scope).
    private const int MaxSimpleUploadBytes = 4_194_304;
    private const int MaxDownloadBytes = 2_097_152;
    private const int MaxTextChars = 20000;

    private readonly GraphServiceClient _client;
    private string? _driveId;
    private string? _rootId;

    public GraphDriveService(GraphServiceClient client, IOptions<GraphAuthOptions> options)
    {
        // OneDrive via /me/drive requires a signed-in user. App-only would
        // need Sites.Selected + a different path scheme (out of scope).
        if (options.Value.AuthMode == AuthMode.AppOnly)
        {
            throw GraphServiceException.AuthMisconfigured(
                "The OneDrive host requires delegated auth. Set Graph:AuthMode to Delegated.");
        }

        _client = client;
    }

    public async Task<DriveInfoDto> GetDriveAsync(CancellationToken ct = default)
    {
        var drive = await _client.Me.Drive.GetAsync(cancellationToken: ct).ConfigureAwait(false);
        if (drive?.Id is null)
        {
            throw GraphServiceException.MailboxUnavailable(null, "Default drive has no id.");
        }

        _driveId = drive.Id;
        return DriveMapper.MapDrive(drive);
    }

    public async Task<IReadOnlyList<DriveInfoDto>> ListDrivesAsync(CancellationToken ct = default)
    {
        var page = await _client.Me.Drives.GetAsync(c =>
            c.QueryParameters.Select = ["id", "name", "driveType", "quota"], ct).ConfigureAwait(false);
        return [.. (page?.Value ?? []).Select(DriveMapper.MapDrive)];
    }

    public async Task<DriveItemSummary> GetItemAsync(string itemRef, CancellationToken ct = default)
    {
        var item = await GetItemOrNullAsync(itemRef, ct).ConfigureAwait(false);
        return item is null
            ? throw GraphServiceException.DriveItemNotFound(itemRef, "onedrive_get_item")
            : DriveMapper.MapItem(item);
    }

    public async Task<IReadOnlyList<DriveItemSummary>> ListChildrenAsync(
        string folderRef = "root", int top = 50, CancellationToken ct = default)
    {
        int take = Math.Clamp(top, 1, 200);
        var builder = await ItemBuilderAsync(folderRef, ct).ConfigureAwait(false);
        var page = await builder.Children.GetAsync(c =>
        {
            c.QueryParameters.Top = take;
            c.QueryParameters.Select = ItemSelect;
            c.QueryParameters.Orderby = ["folder,name"];
        }, ct).ConfigureAwait(false);
        return [.. (page?.Value ?? []).Select(DriveMapper.MapItem)];
    }

    public async Task<IReadOnlyList<DriveItemSummary>> SearchAsync(
        string query, int top = 25, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            throw GraphServiceException.InvalidRequest(
                "Query must not be empty.",
                "pass a filename or keyword, e.g. \"Rechnung\"");
        }

        int take = Math.Clamp(top, 1, 50);
        string driveId = await DriveIdAsync(ct).ConfigureAwait(false);
        var page = await _client.Drives[driveId].Items["root"].SearchWithQ(query.Trim()).GetAsSearchWithQGetResponseAsync(c =>
        {
            c.QueryParameters.Top = take;
            c.QueryParameters.Select = ItemSelect;
        }, ct).ConfigureAwait(false);
        return [.. (page?.Value ?? []).Select(DriveMapper.MapItem)];
    }

    public async Task<FileContentDto> DownloadAsync(
        string itemRef, int maxBytes = 786432, CancellationToken ct = default)
    {
        int cap = Math.Clamp(maxBytes, 1, MaxDownloadBytes);
        var item = await GetItemOrNullAsync(itemRef, ct).ConfigureAwait(false)
            ?? throw GraphServiceException.DriveItemNotFound(itemRef, "onedrive_download_file");

        if (item.Folder is not null)
        {
            throw GraphServiceException.InvalidRequest(
                $"Drive item '{item.Name}' is a folder.",
                "download files only; use onedrive_list_children to browse folders");
        }

        long size = item.Size ?? 0;
        if (size > cap)
        {
            throw GraphServiceException.AttachmentTooLarge(item.Name ?? "?", (int)Math.Min(size, int.MaxValue), cap);
        }

        string driveId = await DriveIdAsync(ct).ConfigureAwait(false);
        var builder = ItemBuilderFor(driveId, await RootIdAsync(ct).ConfigureAwait(false), itemRef);
        await using var stream = await builder.Content.GetAsync(cancellationToken: ct).ConfigureAwait(false)
            ?? throw GraphServiceException.GraphError(0, null, "Download returned no content.");
        using var ms = new MemoryStream();
        await stream.CopyToAsync(ms, ct).ConfigureAwait(false);
        byte[] bytes = ms.ToArray();

        if (DriveMapper.IsTextContent(item.File?.MimeType))
        {
            string text = System.Text.Encoding.UTF8.GetString(bytes);
            bool truncated = text.Length > MaxTextChars;
            return new FileContentDto(
                item.Id ?? string.Empty, item.Name ?? string.Empty, item.File?.MimeType,
                bytes.Length, "text",
                truncated ? DriveMapper.Truncate(text, MaxTextChars) : text,
                null, truncated);
        }

        return new FileContentDto(
            item.Id ?? string.Empty, item.Name ?? string.Empty, item.File?.MimeType,
            bytes.Length, "base64", null, Convert.ToBase64String(bytes), false);
    }

    public async Task<DriveItemSummary> CreateFolderAsync(
        string name, string parentRef = "root", CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(name) || name.IndexOfAny(['/', '\\']) >= 0)
        {
            throw GraphServiceException.InvalidRequest(
                "Folder name must be non-empty and contain no slashes.",
                "pass a plain name, e.g. \"Belege\"");
        }

        string driveId = await DriveIdAsync(ct).ConfigureAwait(false);
        string parentId = await ResolveIdAsync(parentRef, ct).ConfigureAwait(false);
        var folder = new DriveItem
        {
            Name = name.Trim(),
            Folder = new Folder(),
            AdditionalData = new Dictionary<string, object>
            {
                ["@microsoft.graph.conflictBehavior"] = "rename"
            }
        };
        var created = await _client.Drives[driveId].Items[parentId].Children
            .PostAsync(folder, cancellationToken: ct).ConfigureAwait(false);
        return created is null
            ? throw GraphServiceException.GraphError(0, null, "Folder creation returned no result.")
            : DriveMapper.MapItem(created);
    }

    public async Task<DriveItemSummary> UploadAsync(
        string fileName,
        string? parentRef = null,
        string? contentText = null,
        string? contentBase64 = null,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(fileName) || fileName.IndexOfAny(['/', '\\']) >= 0)
        {
            throw GraphServiceException.InvalidRequest(
                "File name must be non-empty and contain no slashes.",
                "pass a plain file name, e.g. \"notiz.txt\"");
        }

        byte[] bytes;
        try
        {
            bytes = contentText is not null
                ? System.Text.Encoding.UTF8.GetBytes(contentText)
                : contentBase64 is not null
                    ? Convert.FromBase64String(contentBase64)
                    : throw GraphServiceException.InvalidRequest(
                        "Either contentText or contentBase64 is required.",
                        "pass text directly or base64 for binary files");
        }
        catch (FormatException)
        {
            throw GraphServiceException.InvalidRequest(
                "contentBase64 is not valid base64.",
                "pass correctly padded base64, or use contentText for text files");
        }

        if (bytes.Length > MaxSimpleUploadBytes)
        {
            throw GraphServiceException.InvalidRequest(
                $"File is {bytes.Length} bytes, above the {MaxSimpleUploadBytes} byte simple-upload limit.",
                "split the file or upload it via OneDrive/SharePoint UI (resumable sessions are out of scope)");
        }

        string driveId = await DriveIdAsync(ct).ConfigureAwait(false);
        string parentId = await ResolveIdAsync(parentRef ?? "root", ct).ConfigureAwait(false);
        using var stream = new MemoryStream(bytes);
        var created = await _client.Drives[driveId].Items[parentId].ItemWithPath(fileName.Trim()).Content
            .PutAsync(stream, cancellationToken: ct).ConfigureAwait(false);
        return created is null
            ? throw GraphServiceException.GraphError(0, null, "Upload returned no result.")
            : DriveMapper.MapItem(created);
    }

    public async Task<DriveItemSummary> MoveAsync(
        string itemId,
        string? newParentRef = null,
        string? newName = null,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(itemId))
        {
            throw GraphServiceException.MissingRef("itemId");
        }

        if (newParentRef is null && newName is null)
        {
            throw GraphServiceException.InvalidRequest(
                "Nothing to do: provide newParentRef and/or newName.",
                "pass a destination folder and/or a new file name");
        }

        string driveId = await DriveIdAsync(ct).ConfigureAwait(false);
        string? parentId = newParentRef is null ? null : await ResolveIdAsync(newParentRef, ct).ConfigureAwait(false);
        var patch = new DriveItem
        {
            Name = newName?.Trim(),
            ParentReference = parentId is null ? null : new ItemReference { Id = parentId }
        };
        var moved = await _client.Drives[driveId].Items[itemId.Trim()].PatchAsync(patch, cancellationToken: ct)
            .ConfigureAwait(false);
        return moved is null
            ? throw GraphServiceException.DriveItemNotFound(itemId, "onedrive_move_item")
            : DriveMapper.MapItem(moved);
    }

    private async Task<string> DriveIdAsync(CancellationToken ct) =>
        _driveId ??= (await _client.Me.Drive.GetAsync(c =>
            c.QueryParameters.Select = ["id"], ct).ConfigureAwait(false))?.Id
            ?? throw GraphServiceException.GraphError(0, null, "Default drive returned no id.");

    private async Task<string> RootIdAsync(CancellationToken ct) =>
        _rootId ??= (await _client.Drives[await DriveIdAsync(ct).ConfigureAwait(false)].Root
            .GetAsync(c => c.QueryParameters.Select = ["id"], ct).ConfigureAwait(false))?.Id
            ?? throw GraphServiceException.GraphError(0, null, "Drive root returned no id.");

    private async Task<string> ResolveIdAsync(string itemRef, CancellationToken ct)
    {
        string reference = PathResolver.RequireRef(itemRef);
        if (PathResolver.IsRoot(reference))
        {
            return await RootIdAsync(ct).ConfigureAwait(false);
        }

        if (!PathResolver.IsPath(reference))
        {
            return reference;
        }

        var item = await GetItemOrNullAsync(reference, ct).ConfigureAwait(false);
        return item?.Id
            ?? throw GraphServiceException.DriveItemNotFound(reference, "resolve");
    }

    private async Task<DriveItem?> GetItemOrNullAsync(string itemRef, CancellationToken ct)
    {
        string reference = PathResolver.RequireRef(itemRef);
        string driveId = await DriveIdAsync(ct).ConfigureAwait(false);
        return await ItemBuilderFor(driveId, await RootIdAsync(ct).ConfigureAwait(false), reference)
            .GetAsync(c => c.QueryParameters.Select = ItemSelect, ct).ConfigureAwait(false);
    }

    private Microsoft.Graph.Drives.Item.Items.Item.DriveItemItemRequestBuilder ItemBuilderFor(
        string driveId, string rootId, string reference)
    {
        if (PathResolver.IsRoot(reference))
        {
            return _client.Drives[driveId].Items[rootId];
        }

        return PathResolver.IsPath(reference)
            ? _client.Drives[driveId].Root.ItemWithPath(PathResolver.NormalizePath(reference))
            : _client.Drives[driveId].Items[reference.Trim()];
    }

    private async Task<Microsoft.Graph.Drives.Item.Items.Item.DriveItemItemRequestBuilder> ItemBuilderAsync(
        string itemRef, CancellationToken ct) =>
        ItemBuilderFor(await DriveIdAsync(ct).ConfigureAwait(false),
            await RootIdAsync(ct).ConfigureAwait(false),
            PathResolver.RequireRef(itemRef));
}
