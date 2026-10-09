using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;
using System.IO.Abstractions;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;
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
    // use a resumable upload session, see UploadViaSessionAsync).
    private const int MaxSimpleUploadBytes = 4_194_304;
    // Resumable sessions support up to 250 GB API-side; the host caps at
    // 100 MiB so a single tool call cannot run for very long.
    private const long MaxResumableUploadBytes = 104_857_600;
    // Fragment size: Graph requires multiples of 320 KiB, recommends 5-10 MiB.
    private const int UploadChunkSize = 5_242_880; // 5 MiB = 16 x 320 KiB
    // Children are paged at the Graph maximum; the client-side scan stops here
    // so one tool call cannot page a pathological folder forever.
    private const int PageSize = 200;
    private const int MaxScannedChildren = 5000;
    private const int MaxDownloadBytes = 2_097_152;
    private const int MaxTextChars = 20000;

    private readonly GraphServiceClient _client;
    private readonly HttpClient _uploadHttp;
    private readonly IFileSystem _fileSystem;
    private static readonly HttpClient SharedUploadHttp = new();
    private string? _driveId;
    private string? _rootId;

    public GraphDriveService(
        GraphServiceClient client,
        IOptions<GraphAuthOptions> options,
        HttpClient? uploadHttp = null,
        IFileSystem? fileSystem = null)
    {
        // OneDrive via /me/drive requires a signed-in user. App-only would
        // need Sites.Selected + a different path scheme (out of scope).
        if (options.Value.AuthMode == AuthMode.AppOnly)
        {
            throw GraphServiceException.AuthMisconfigured(
                "The OneDrive host requires delegated auth. Set Graph:AuthMode to Delegated.");
        }

        _client = client;
        // Upload-session fragment PUTs go to a preauthenticated upload URL and
        // must NOT carry the Graph Authorization header, hence a separate client.
        _uploadHttp = uploadHttp ?? SharedUploadHttp;
        _fileSystem = fileSystem ?? new FileSystem();
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
        // NOTE: Graph supports $orderby only for name/size/lastModifiedDateTime
        // ("folder" would yield 400), and $top applies before any client-side
        // sort. Page the full candidate set first, then sort folders-first and
        // apply Take so later-named folders are not silently omitted.
        var children = new List<DriveItem>();
        var page = await builder.Children.GetAsync(c =>
        {
            c.QueryParameters.Top = PageSize;
            c.QueryParameters.Select = ItemSelect;
            c.QueryParameters.Orderby = ["name"];
        }, ct).ConfigureAwait(false);
        while (page is not null && children.Count < MaxScannedChildren)
        {
            children.AddRange(page.Value ?? []);
            if (string.IsNullOrWhiteSpace(page.OdataNextLink) || children.Count >= MaxScannedChildren)
            {
                break;
            }

            page = await builder.Children.WithUrl(page.OdataNextLink)
                .GetAsync(cancellationToken: ct).ConfigureAwait(false);
        }

        return [.. children.Select(DriveMapper.MapItem)
            .OrderByDescending(i => i.IsFolder)
            .ThenBy(i => i.Name, StringComparer.OrdinalIgnoreCase)
            .Take(take)];
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

    public async Task<FileContentDto> DownloadToFileAsync(
        string itemRef, string localPath, bool overwrite = false, CancellationToken ct = default)
    {
        string dest = PathResolver.RequireAbsolutePath(localPath);
        if (_fileSystem.File.Exists(dest) && !overwrite)
        {
            throw GraphServiceException.InvalidRequest(
                $"Local file '{dest}' already exists.",
                "pass overwrite:true to replace it, or choose another localPath");
        }

        var item = await GetItemOrNullAsync(itemRef, ct).ConfigureAwait(false)
            ?? throw GraphServiceException.DriveItemNotFound(itemRef, "onedrive_download_file");

        if (item.Folder is not null)
        {
            throw GraphServiceException.InvalidRequest(
                $"Drive item '{item.Name}' is a folder.",
                "download files only; use onedrive_list_children to browse folders");
        }

        string? directory = _fileSystem.Path.GetDirectoryName(dest);
        if (directory is not null)
        {
            _fileSystem.Directory.CreateDirectory(directory);
        }

        // Stream into an adjacent temp file, then move atomically: the Exists
        // check above cannot protect a direct File.Create against concurrent
        // creators, and a failed/cancelled copy must not leave a partial dest.
        string tempPath = $"{dest}.{Guid.NewGuid():N}.tmp";
        try
        {
            string driveId = await DriveIdAsync(ct).ConfigureAwait(false);
            var builder = ItemBuilderFor(driveId, await RootIdAsync(ct).ConfigureAwait(false), itemRef);
            await using var source = await builder.Content.GetAsync(cancellationToken: ct).ConfigureAwait(false)
                ?? throw GraphServiceException.GraphError(0, null, "Download returned no content.");
            await using (var target = _fileSystem.File.Open(
                tempPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                await source.CopyToAsync(target, ct).ConfigureAwait(false);
            }

            _fileSystem.File.Move(tempPath, dest, overwrite);
        }
        catch
        {
            TryDelete(tempPath);
            throw;
        }

        long size = _fileSystem.FileInfo.New(dest).Length;
        return new FileContentDto(
            item.Id ?? string.Empty, item.Name ?? string.Empty, item.File?.MimeType,
            size, "file", null, null, false, dest);
    }

    private void TryDelete(string path)
    {
        try
        {
            if (_fileSystem.File.Exists(path))
            {
                _fileSystem.File.Delete(path);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
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
        string? localPath = null,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(fileName) || fileName.IndexOfAny(['/', '\\']) >= 0)
        {
            throw GraphServiceException.InvalidRequest(
                "File name must be non-empty and contain no slashes.",
                "pass a plain file name, e.g. \"notiz.txt\"");
        }

        if (localPath is not null)
        {
            if (contentText is not null || contentBase64 is not null)
            {
                throw GraphServiceException.InvalidRequest(
                    "Provide either localPath or inline content, not both.",
                    "use localPath for files on disk, contentText for small snippets");
            }

            return await UploadFromFileAsync(fileName, parentRef ?? "root", localPath, ct)
                .ConfigureAwait(false);
        }

        byte[] bytes;
        try
        {
            bytes = contentText is not null
                ? System.Text.Encoding.UTF8.GetBytes(contentText)
                : contentBase64 is not null
                    ? Convert.FromBase64String(contentBase64)
                    : throw GraphServiceException.InvalidRequest(
                        "Either contentText, contentBase64 or localPath is required.",
                        "pass text directly, base64 for small binaries, or a localPath for files on disk");
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
                $"Inline content is {bytes.Length} bytes, above the {MaxSimpleUploadBytes} byte simple-upload limit.",
                "write the content to a file and pass its absolute localPath (resumable upload)");
        }

        string driveId = await DriveIdAsync(ct).ConfigureAwait(false);
        string parentId = await ResolveIdAsync(parentRef ?? "root", ct).ConfigureAwait(false);
        // Simple PUT replaces by default; resolve a free name first so inline
        // uploads follow the same rename-on-conflict policy as the rest.
        string uniqueName = await ResolveUniqueNameAsync(driveId, parentId, fileName.Trim(), ct)
            .ConfigureAwait(false);
        using var stream = new MemoryStream(bytes);
        var created = await _client.Drives[driveId].Items[parentId].ItemWithPath(uniqueName).Content
            .PutAsync(stream, cancellationToken: ct).ConfigureAwait(false);
        return created is null
            ? throw GraphServiceException.GraphError(0, null, "Upload returned no result.")
            : DriveMapper.MapItem(created);
    }

    /// <summary>Uploads a host file. Bytes are read from disk (never through
    /// the model context): simple PUT up to 4 MiB, resumable session above.</summary>
    private async Task<DriveItemSummary> UploadFromFileAsync(
        string fileName, string parentRef, string localPath, CancellationToken ct)
    {
        string full = PathResolver.RequireAbsolutePath(localPath);
        var info = _fileSystem.FileInfo.New(full);
        if (!info.Exists)
        {
            throw GraphServiceException.InvalidRequest(
                $"Local file '{full}' does not exist.",
                "pass the absolute path of an existing file on the host");
        }

        if (info.Length > MaxResumableUploadBytes)
        {
            throw GraphServiceException.InvalidRequest(
                $"Local file is {info.Length} bytes, above the {MaxResumableUploadBytes} byte tool limit.",
                "upload it via the OneDrive UI (API sessions support up to 250 GB)");
        }

        string driveId = await DriveIdAsync(ct).ConfigureAwait(false);
        string parentId = await ResolveIdAsync(parentRef, ct).ConfigureAwait(false);

        if (info.Length <= MaxSimpleUploadBytes)
        {
            string uniqueName = await ResolveUniqueNameAsync(driveId, parentId, fileName.Trim(), ct)
                .ConfigureAwait(false);
            await using var stream = _fileSystem.File.OpenRead(full);
            var created = await _client.Drives[driveId].Items[parentId].ItemWithPath(uniqueName).Content
                .PutAsync(stream, cancellationToken: ct).ConfigureAwait(false);
            return created is null
                ? throw GraphServiceException.GraphError(0, null, "Upload returned no result.")
                : DriveMapper.MapItem(created);
        }

        return await UploadViaSessionAsync(driveId, parentId, fileName.Trim(), full, info.Length, ct)
            .ConfigureAwait(false);
    }

    /// <summary>Resumable upload (Graph v1.0 createUploadSession): sequential
    /// fragments, each a multiple of 320 KiB, PUT to the preauthenticated URL.</summary>
    private async Task<DriveItemSummary> UploadViaSessionAsync(
        string driveId, string parentId, string fileName,
        string localPath, long totalBytes, CancellationToken ct)
    {
        var session = await _client.Drives[driveId].Items[parentId].ItemWithPath(fileName)
            .CreateUploadSession.PostAsync(new CreateUploadSessionPostRequestBody
            {
                Item = new DriveItemUploadableProperties
                {
                    AdditionalData = new Dictionary<string, object>
                    {
                        ["@microsoft.graph.conflictBehavior"] = "rename"
                    }
                }
            }, cancellationToken: ct).ConfigureAwait(false);
        string uploadUrl = session?.UploadUrl
            ?? throw GraphServiceException.GraphError(0, null, "Upload session returned no upload URL.");

        await using var file = _fileSystem.File.OpenRead(localPath);
        byte[] buffer = new byte[UploadChunkSize];
        long offset = 0;
        while (offset < totalBytes)
        {
            int read = 0;
            while (read < buffer.Length)
            {
                int n = await file.ReadAsync(buffer.AsMemory(read), ct).ConfigureAwait(false);
                if (n == 0)
                {
                    break;
                }

                read += n;
            }

            if (read == 0)
            {
                throw GraphServiceException.GraphError(
                    0, null, "Local file shrank during upload; retry the upload.");
            }

            using var fragment = new ByteArrayContent(buffer, 0, read);
            fragment.Headers.ContentRange = new ContentRangeHeaderValue(offset, offset + read - 1, totalBytes);
            using var response = await _uploadHttp.PutAsync(uploadUrl, fragment, ct).ConfigureAwait(false);

            if (response.StatusCode is HttpStatusCode.OK or HttpStatusCode.Created)
            {
                string json = await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
                return MapSessionResult(json);
            }

            if (response.StatusCode != HttpStatusCode.Accepted)
            {
                string body = (await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false)).Trim();
                throw GraphServiceException.GraphError(
                    (int)response.StatusCode, null,
                    $"Upload fragment {offset}-{offset + read - 1} failed{(body.Length > 0 ? $": {body[..Math.Min(body.Length, 300)]}" : ".")}");
            }

            offset += read;
        }

        throw GraphServiceException.GraphError(0, null, "Upload session completed without a final result.");
    }

    /// <summary>Rename-on-conflict for simple PUTs (which replace by
    /// default): returns <paramref name="fileName"/> when free, else
    /// "stem 1.ext", "stem 2.ext", … – the same outcome the resumable branch
    /// gets server-side via conflictBehavior=rename.</summary>
    private async Task<string> ResolveUniqueNameAsync(
        string driveId, string parentId, string fileName, CancellationToken ct)
    {
        var taken = await ListChildNamesAsync(driveId, parentId, ct).ConfigureAwait(false);
        if (!taken.Contains(fileName))
        {
            return fileName;
        }

        string stem = Path.GetFileNameWithoutExtension(fileName);
        string ext = Path.GetExtension(fileName);
        for (int i = 1; ; i++)
        {
            string candidate = $"{stem} {i}{ext}";
            if (!taken.Contains(candidate))
            {
                return candidate;
            }
        }
    }

    private async Task<HashSet<string>> ListChildNamesAsync(
        string driveId, string parentId, CancellationToken ct)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var page = await _client.Drives[driveId].Items[parentId].Children.GetAsync(c =>
        {
            c.QueryParameters.Top = PageSize;
            c.QueryParameters.Select = ["name"];
        }, ct).ConfigureAwait(false);
        while (page is not null && names.Count < MaxScannedChildren)
        {
            foreach (var child in page.Value ?? [])
            {
                if (child.Name is not null)
                {
                    names.Add(child.Name);
                }
            }

            if (string.IsNullOrWhiteSpace(page.OdataNextLink) || names.Count >= MaxScannedChildren)
            {
                break;
            }

            page = await _client.Drives[driveId].Items[parentId].Children
                .WithUrl(page.OdataNextLink)
                .GetAsync(cancellationToken: ct).ConfigureAwait(false);
        }

        return names;
    }

    private static DriveItemSummary MapSessionResult(string json)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            string id = root.TryGetProperty("id", out var idProp)
                ? idProp.GetString() ?? string.Empty : string.Empty;
            string name = root.TryGetProperty("name", out var nameProp)
                ? nameProp.GetString() ?? string.Empty : string.Empty;
            long size = root.TryGetProperty("size", out var sizeProp) && sizeProp.TryGetInt64(out long s)
                ? s : 0;
            string? mime = root.TryGetProperty("file", out var fileProp)
                && fileProp.TryGetProperty("mimeType", out var mimeProp)
                ? mimeProp.GetString() : null;
            string? parentId = root.TryGetProperty("parentReference", out var parentProp)
                && parentProp.TryGetProperty("id", out var parentIdProp)
                ? parentIdProp.GetString() : null;
            string? webUrl = root.TryGetProperty("webUrl", out var webProp)
                ? webProp.GetString() : null;
            DateTimeOffset? modified = root.TryGetProperty("lastModifiedDateTime", out var modProp)
                && modProp.TryGetDateTimeOffset(out var dto) ? dto : null;
            return new DriveItemSummary(id, name, false, size, mime, parentId, webUrl, modified, 0);
        }
        catch (JsonException ex)
        {
            throw GraphServiceException.GraphError(0, null, $"Upload finished but the result was not JSON: {ex.Message}");
        }
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
