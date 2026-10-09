using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;
using System.IO.Abstractions;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;
using Microsoft.Graph.Models;
using MicrosoftMcp.Common;

namespace MicrosoftMcp.SharePoint;

/// <summary>SharePoint sites + file access on explicit drives (/sites,
/// /drives/{driveId}). Read, create and move only – no delete. Items are
/// addressed by "root", id or /path, always scoped to the given driveId.
/// Site ids come from sharepoint_search_sites ("root" = tenant root site),
/// drive ids from sharepoint_list_site_drives or
/// teams_get_channel_files_folder.</summary>
public sealed class GraphSharePointService : IGraphSharePointService
{
    private static readonly string[] SiteSelect = ["id", "name", "displayName", "webUrl"];
    private static readonly string[] DriveSelect = ["id", "name", "driveType", "quota", "webUrl"];
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
    private readonly Dictionary<string, string> _rootIds = new(StringComparer.OrdinalIgnoreCase);

    public GraphSharePointService(
        GraphServiceClient client,
        IOptions<GraphAuthOptions> options,
        HttpClient? uploadHttp = null,
        IFileSystem? fileSystem = null)
    {
        // SharePoint via /sites/* requires a signed-in user. App-only would
        // need Sites.Selected + a different path scheme (out of scope).
        if (options.Value.AuthMode == AuthMode.AppOnly)
        {
            throw GraphServiceException.AuthMisconfigured(
                "The SharePoint host requires delegated auth. Set Graph:AuthMode to Delegated.");
        }

        _client = client;
        // Upload-session fragment PUTs go to a preauthenticated upload URL and
        // must NOT carry the Graph Authorization header, hence a separate client.
        _uploadHttp = uploadHttp ?? SharedUploadHttp;
        _fileSystem = fileSystem ?? new FileSystem();
    }

    public async Task<IReadOnlyList<SharePointSiteInfo>> SearchSitesAsync(
        string query, int top = 25, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            throw GraphServiceException.InvalidRequest(
                "Query must not be empty.",
                "pass a site or team name, e.g. \"Engineering\" (or \"root\" in sharepoint_get_site for the tenant root site)");
        }

        int take = Math.Clamp(top, 1, 50);
        var page = await _client.Sites.GetAsync(c =>
        {
            c.QueryParameters.Search = query.Trim();
            c.QueryParameters.Top = take;
            c.QueryParameters.Select = SiteSelect;
        }, ct).ConfigureAwait(false);
        return [.. (page?.Value ?? []).Select(SharePointMapper.MapSite).Take(take)];
    }

    public async Task<SharePointSiteInfo> GetSiteAsync(string siteId, CancellationToken ct = default)
    {
        string id = PathResolver.RequireSiteId(siteId);
        var site = await _client.Sites[id].GetAsync(c =>
        {
            c.QueryParameters.Select = SiteSelect;
        }, ct).ConfigureAwait(false);
        return site?.Id is null
            ? throw GraphServiceException.SiteNotFound(siteId, "sharepoint_get_site")
            : SharePointMapper.MapSite(site);
    }

    public async Task<IReadOnlyList<SharePointDriveInfo>> ListSiteDrivesAsync(
        string siteId, int top = 50, CancellationToken ct = default)
    {
        string id = PathResolver.RequireSiteId(siteId);
        int take = Math.Clamp(top, 1, 100);
        var drives = new List<Drive>();
        var page = await _client.Sites[id].Drives.GetAsync(c =>
        {
            c.QueryParameters.Top = Math.Min(take, PageSize);
            c.QueryParameters.Select = DriveSelect;
        }, ct).ConfigureAwait(false);
        while (page is not null && drives.Count < take)
        {
            drives.AddRange(page.Value ?? []);
            if (string.IsNullOrWhiteSpace(page.OdataNextLink) || drives.Count >= take)
            {
                break;
            }

            page = await _client.Sites[id].Drives
                .WithUrl(page.OdataNextLink)
                .GetAsync(cancellationToken: ct).ConfigureAwait(false);
        }

        return [.. drives.Select(SharePointMapper.MapDrive).Take(take)];
    }

    public async Task<DriveItemSummary> GetItemAsync(
        string driveId, string itemRef, CancellationToken ct = default)
    {
        string drive = PathResolver.RequireDriveId(driveId);
        var item = await GetItemOrNullAsync(drive, itemRef, ct).ConfigureAwait(false);
        return item is null
            ? throw GraphServiceException.SharePointItemNotFound(itemRef, "sharepoint_get_item")
            : SharePointMapper.MapItem(item);
    }

    public async Task<IReadOnlyList<DriveItemSummary>> ListChildrenAsync(
        string driveId, string folderRef = "root", int top = 50, CancellationToken ct = default)
    {
        string drive = PathResolver.RequireDriveId(driveId);
        int take = Math.Clamp(top, 1, 200);
        var builder = await ItemBuilderAsync(drive, folderRef, ct).ConfigureAwait(false);
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

        return [.. children.Select(SharePointMapper.MapItem)
            .OrderByDescending(i => i.IsFolder)
            .ThenBy(i => i.Name, StringComparer.OrdinalIgnoreCase)
            .Take(take)];
    }

    public async Task<IReadOnlyList<DriveItemSummary>> SearchAsync(
        string driveId, string query, int top = 25, CancellationToken ct = default)
    {
        string drive = PathResolver.RequireDriveId(driveId);
        if (string.IsNullOrWhiteSpace(query))
        {
            throw GraphServiceException.InvalidRequest(
                "Query must not be empty.",
                "pass a filename or keyword, e.g. \"Rechnung\"");
        }

        int take = Math.Clamp(top, 1, 50);
        var page = await _client.Drives[drive].Items["root"].SearchWithQ(query.Trim()).GetAsSearchWithQGetResponseAsync(c =>
        {
            c.QueryParameters.Top = take;
            c.QueryParameters.Select = ItemSelect;
        }, ct).ConfigureAwait(false);
        return [.. (page?.Value ?? []).Select(SharePointMapper.MapItem)];
    }

    public async Task<FileContentDto> DownloadAsync(
        string driveId, string itemRef, int maxBytes = 786432, CancellationToken ct = default)
    {
        string drive = PathResolver.RequireDriveId(driveId);
        int cap = Math.Clamp(maxBytes, 1, MaxDownloadBytes);
        var item = await GetItemOrNullAsync(drive, itemRef, ct).ConfigureAwait(false)
            ?? throw GraphServiceException.SharePointItemNotFound(itemRef, "sharepoint_download_file");

        if (item.Folder is not null)
        {
            throw GraphServiceException.InvalidRequest(
                $"Drive item '{item.Name}' is a folder.",
                "download files only; use sharepoint_list_children to browse folders");
        }

        long size = item.Size ?? 0;
        if (size > cap)
        {
            throw GraphServiceException.AttachmentTooLarge(item.Name ?? "?", (int)Math.Min(size, int.MaxValue), cap);
        }

        var builder = ItemBuilderFor(drive, await RootIdAsync(drive, ct).ConfigureAwait(false), itemRef);
        await using var stream = await builder.Content.GetAsync(cancellationToken: ct).ConfigureAwait(false)
            ?? throw GraphServiceException.GraphError(0, null, "Download returned no content.");
        using var ms = new MemoryStream();
        await stream.CopyToAsync(ms, ct).ConfigureAwait(false);
        byte[] bytes = ms.ToArray();

        if (SharePointMapper.IsTextContent(item.File?.MimeType))
        {
            string text = System.Text.Encoding.UTF8.GetString(bytes);
            bool truncated = text.Length > MaxTextChars;
            return new FileContentDto(
                item.Id ?? string.Empty, item.Name ?? string.Empty, item.File?.MimeType,
                bytes.Length, "text",
                truncated ? SharePointMapper.Truncate(text, MaxTextChars) : text,
                null, truncated);
        }

        return new FileContentDto(
            item.Id ?? string.Empty, item.Name ?? string.Empty, item.File?.MimeType,
            bytes.Length, "base64", null, Convert.ToBase64String(bytes), false);
    }

    public async Task<FileContentDto> DownloadToFileAsync(
        string driveId, string itemRef, string localPath, bool overwrite = false, CancellationToken ct = default)
    {
        string drive = PathResolver.RequireDriveId(driveId);
        string dest = PathResolver.RequireAbsolutePath(localPath);
        if (_fileSystem.File.Exists(dest) && !overwrite)
        {
            throw GraphServiceException.InvalidRequest(
                $"Local file '{dest}' already exists.",
                "pass overwrite:true to replace it, or choose another localPath");
        }

        var item = await GetItemOrNullAsync(drive, itemRef, ct).ConfigureAwait(false)
            ?? throw GraphServiceException.SharePointItemNotFound(itemRef, "sharepoint_download_file");

        if (item.Folder is not null)
        {
            throw GraphServiceException.InvalidRequest(
                $"Drive item '{item.Name}' is a folder.",
                "download files only; use sharepoint_list_children to browse folders");
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
            var builder = ItemBuilderFor(drive, await RootIdAsync(drive, ct).ConfigureAwait(false), itemRef);
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
        string driveId, string name, string parentRef = "root", CancellationToken ct = default)
    {
        string drive = PathResolver.RequireDriveId(driveId);
        if (string.IsNullOrWhiteSpace(name) || name.IndexOfAny(['/', '\\']) >= 0)
        {
            throw GraphServiceException.InvalidRequest(
                "Folder name must be non-empty and contain no slashes.",
                "pass a plain name, e.g. \"Belege\"");
        }

        string parentId = await ResolveIdAsync(drive, parentRef, ct).ConfigureAwait(false);
        var folder = new DriveItem
        {
            Name = name.Trim(),
            Folder = new Folder(),
            AdditionalData = new Dictionary<string, object>
            {
                ["@microsoft.graph.conflictBehavior"] = "rename"
            }
        };
        var created = await _client.Drives[drive].Items[parentId].Children
            .PostAsync(folder, cancellationToken: ct).ConfigureAwait(false);
        return created is null
            ? throw GraphServiceException.GraphError(0, null, "Folder creation returned no result.")
            : SharePointMapper.MapItem(created);
    }

    public async Task<DriveItemSummary> UploadAsync(
        string driveId,
        string fileName,
        string? parentRef = null,
        string? contentText = null,
        string? contentBase64 = null,
        string? localPath = null,
        CancellationToken ct = default)
    {
        string drive = PathResolver.RequireDriveId(driveId);
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

            return await UploadFromFileAsync(drive, fileName, parentRef ?? "root", localPath, ct)
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

        string parentId = await ResolveIdAsync(drive, parentRef ?? "root", ct).ConfigureAwait(false);
        // Simple PUT replaces by default; resolve a free name first so inline
        // uploads follow the same rename-on-conflict policy as the rest.
        string uniqueName = await ResolveUniqueNameAsync(drive, parentId, fileName.Trim(), ct)
            .ConfigureAwait(false);
        using var stream = new MemoryStream(bytes);
        var created = await _client.Drives[drive].Items[parentId].ItemWithPath(uniqueName).Content
            .PutAsync(stream, cancellationToken: ct).ConfigureAwait(false);
        return created is null
            ? throw GraphServiceException.GraphError(0, null, "Upload returned no result.")
            : SharePointMapper.MapItem(created);
    }

    /// <summary>Uploads a host file. Bytes are read from disk (never through
    /// the model context): simple PUT up to 4 MiB, resumable session above.</summary>
    private async Task<DriveItemSummary> UploadFromFileAsync(
        string driveId, string fileName, string parentRef, string localPath, CancellationToken ct)
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
                "upload it via the SharePoint UI (API sessions support up to 250 GB)");
        }

        string parentId = await ResolveIdAsync(driveId, parentRef, ct).ConfigureAwait(false);

        if (info.Length <= MaxSimpleUploadBytes)
        {
            string uniqueName = await ResolveUniqueNameAsync(driveId, parentId, fileName.Trim(), ct)
                .ConfigureAwait(false);
            await using var stream = _fileSystem.File.OpenRead(full);
            var created = await _client.Drives[driveId].Items[parentId].ItemWithPath(uniqueName).Content
                .PutAsync(stream, cancellationToken: ct).ConfigureAwait(false);
            return created is null
                ? throw GraphServiceException.GraphError(0, null, "Upload returned no result.")
                : SharePointMapper.MapItem(created);
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
        string driveId,
        string itemId,
        string? newParentRef = null,
        string? newName = null,
        CancellationToken ct = default)
    {
        string drive = PathResolver.RequireDriveId(driveId);
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

        string? parentId = newParentRef is null ? null : await ResolveIdAsync(drive, newParentRef, ct).ConfigureAwait(false);
        var patch = new DriveItem
        {
            Name = newName?.Trim(),
            ParentReference = parentId is null ? null : new ItemReference { Id = parentId }
        };
        var moved = await _client.Drives[drive].Items[itemId.Trim()].PatchAsync(patch, cancellationToken: ct)
            .ConfigureAwait(false);
        return moved is null
            ? throw GraphServiceException.SharePointItemNotFound(itemId, "sharepoint_move_item")
            : SharePointMapper.MapItem(moved);
    }

    private async Task<string> RootIdAsync(string driveId, CancellationToken ct)
    {
        if (_rootIds.TryGetValue(driveId, out string? cached))
        {
            return cached;
        }

        string rootId = (await _client.Drives[driveId].Root
            .GetAsync(c => c.QueryParameters.Select = ["id"], ct).ConfigureAwait(false))?.Id
            ?? throw GraphServiceException.GraphError(0, null, "Drive root returned no id.");
        _rootIds[driveId] = rootId;
        return rootId;
    }

    private async Task<string> ResolveIdAsync(string driveId, string itemRef, CancellationToken ct)
    {
        string reference = PathResolver.RequireRef(itemRef);
        if (PathResolver.IsRoot(reference))
        {
            return await RootIdAsync(driveId, ct).ConfigureAwait(false);
        }

        if (!PathResolver.IsPath(reference))
        {
            return reference;
        }

        var item = await GetItemOrNullAsync(driveId, reference, ct).ConfigureAwait(false);
        return item?.Id
            ?? throw GraphServiceException.SharePointItemNotFound(reference, "resolve");
    }

    private async Task<DriveItem?> GetItemOrNullAsync(string driveId, string itemRef, CancellationToken ct)
    {
        string reference = PathResolver.RequireRef(itemRef);
        return await ItemBuilderFor(driveId, await RootIdAsync(driveId, ct).ConfigureAwait(false), reference)
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
        string driveId, string itemRef, CancellationToken ct) =>
        ItemBuilderFor(driveId,
            await RootIdAsync(driveId, ct).ConfigureAwait(false),
            PathResolver.RequireRef(itemRef));
}
