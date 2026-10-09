namespace MicrosoftMcp.SharePoint.Tests;

using System.IO.Abstractions;
using MicrosoftMcp.Common;

/// <summary>In-memory fake of <see cref="IGraphSharePointService"/>. No Graph, no network.</summary>
internal sealed class FakeGraphSharePointService : IGraphSharePointService
{
    private sealed class Node
    {
        public required string Id { get; init; }
        public string Name { get; set; } = string.Empty;
        public bool IsFolder { get; init; }
        public string? ParentId { get; set; }
        public string? MimeType { get; init; }
        public byte[] Content { get; set; } = [];
    }

    private readonly Dictionary<string, SharePointSiteInfo> _sites = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, (string SiteId, SharePointDriveInfo Drive)> _drives = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, Dictionary<string, Node>> _nodes = new(StringComparer.OrdinalIgnoreCase);
    private readonly IFileSystem _fileSystem;
    private int _counter;

    public FakeGraphSharePointService(IFileSystem? fileSystem = null)
    {
        _fileSystem = fileSystem ?? new FileSystem();
        _sites["site-root"] = new SharePointSiteInfo("site-root", "Tenant-Root", "Contoso", "https://contoso.sharepoint.com");
        _sites["site-eng"] = new SharePointSiteInfo("site-eng", "Engineering", "Engineering", "https://contoso.sharepoint.com/sites/eng");
        _drives["drive-eng"] = ("site-eng",
            new SharePointDriveInfo("drive-eng", "Dokumente", "documentLibrary", "https://contoso.sharepoint.com/drive-eng", 10000, 1000, 9000));
        _drives["drive-root"] = ("site-root",
            new SharePointDriveInfo("drive-root", "Freigegeben", "documentLibrary", "https://contoso.sharepoint.com/drive-root", 10000, 500, 9500));

        Add("drive-eng", new Node { Id = "root", Name = "root", IsFolder = true });
        Add("drive-eng", new Node { Id = "f-docs", Name = "Protokolle", IsFolder = true, ParentId = "root" });
        Add("drive-eng", new Node
        {
            Id = "f-note", Name = "notiz.txt", ParentId = "root",
            MimeType = "text/plain", Content = "Hallo Welt\nZeile zwei"u8.ToArray()
        });
        Add("drive-eng", new Node
        {
            Id = "f-img", Name = "bild.png", ParentId = "f-docs",
            MimeType = "image/png", Content = new byte[64]
        });
        Add("drive-eng", new Node
        {
            Id = "f-big", Name = "gross.bin", ParentId = "root",
            MimeType = "application/octet-stream", Content = new byte[3_000_000]
        });
        Add("drive-root", new Node { Id = "root", Name = "root", IsFolder = true });
        Add("drive-root", new Node
        {
            Id = "r-readme", Name = "liesmich.txt", ParentId = "root",
            MimeType = "text/plain", Content = "Root-Laufwerk"u8.ToArray()
        });
    }

    public string? LastSearch { get; private set; }
    public string? LastDrive { get; private set; }

    private void Add(string driveId, Node n)
    {
        if (!_nodes.TryGetValue(driveId, out var nodes))
        {
            nodes = new Dictionary<string, Node>(StringComparer.OrdinalIgnoreCase);
            _nodes[driveId] = nodes;
        }

        nodes[n.Id] = n;
    }

    private Dictionary<string, Node> Drive(string driveId)
    {
        string drive = PathResolver.RequireDriveId(driveId);
        LastDrive = drive;
        return _nodes.TryGetValue(drive, out var nodes)
            ? nodes
            : throw GraphServiceException.SharePointDriveNotFound(driveId, "fake");
    }

    private Node Get(string driveId, string itemRef)
    {
        var nodes = Drive(driveId);
        string reference = PathResolver.RequireRef(itemRef);
        if (PathResolver.IsRoot(reference))
        {
            return nodes["root"];
        }

        if (PathResolver.IsPath(reference))
        {
            string[] parts = PathResolver.NormalizePath(reference).Split('/');
            Node current = nodes["root"];
            foreach (string part in parts)
            {
                current = nodes.Values.FirstOrDefault(n =>
                        string.Equals(n.ParentId, current.Id, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(n.Name, part, StringComparison.OrdinalIgnoreCase))
                    ?? throw GraphServiceException.SharePointItemNotFound(itemRef, "fake");
            }

            return current;
        }

        return nodes.TryGetValue(reference, out var node)
            ? node
            : throw GraphServiceException.SharePointItemNotFound(itemRef, "fake");
    }

    private static DriveItemSummary ToSummary(Node n) => new(
        n.Id, n.Name, n.IsFolder, n.Content.Length,
        n.MimeType, n.ParentId, null, DateTimeOffset.UtcNow,
        n.IsFolder ? 1 : 0);

    public Task<IReadOnlyList<SharePointSiteInfo>> SearchSitesAsync(
        string query, int top = 25, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            throw GraphServiceException.InvalidRequest(
                "Query must not be empty.", "pass a site or team name");
        }

        LastSearch = query;
        int take = Math.Clamp(top, 1, 50);
        return Task.FromResult<IReadOnlyList<SharePointSiteInfo>>(
            [.. _sites.Values
                .Where(s => (s.DisplayName ?? s.Name).Contains(query, StringComparison.OrdinalIgnoreCase))
                .Take(take)]);
    }

    public Task<SharePointSiteInfo> GetSiteAsync(string siteId, CancellationToken ct = default)
    {
        string id = PathResolver.RequireSiteId(siteId);
        if (string.Equals(id, "root", StringComparison.OrdinalIgnoreCase))
        {
            return Task.FromResult(_sites["site-root"]);
        }

        return Task.FromResult(_sites.TryGetValue(id, out var site)
            ? site
            : throw GraphServiceException.SiteNotFound(siteId, "fake"));
    }

    public Task<IReadOnlyList<SharePointDriveInfo>> ListSiteDrivesAsync(
        string siteId, int top = 50, CancellationToken ct = default)
    {
        string id = PathResolver.RequireSiteId(siteId);
        string resolved = string.Equals(id, "root", StringComparison.OrdinalIgnoreCase) ? "site-root" : id;
        if (!_sites.ContainsKey(resolved))
        {
            throw GraphServiceException.SiteNotFound(siteId, "fake");
        }

        int take = Math.Clamp(top, 1, 100);
        return Task.FromResult<IReadOnlyList<SharePointDriveInfo>>(
            [.. _drives.Values.Where(d => d.SiteId == resolved).Select(d => d.Drive).Take(take)]);
    }

    public Task<DriveItemSummary> GetItemAsync(
        string driveId, string itemRef, CancellationToken ct = default) =>
        Task.FromResult(ToSummary(Get(driveId, itemRef)));

    public Task<IReadOnlyList<DriveItemSummary>> ListChildrenAsync(
        string driveId, string folderRef = "root", int top = 50, CancellationToken ct = default)
    {
        var folder = Get(driveId, folderRef);
        if (!folder.IsFolder)
        {
            throw GraphServiceException.InvalidRequest(
                $"Drive item '{folder.Name}' is a folder expected.",
                "pass a folder id or /path");
        }

        int take = Math.Clamp(top, 1, 200);
        var nodes = Drive(driveId);
        return Task.FromResult<IReadOnlyList<DriveItemSummary>>(
            [.. nodes.Values.Where(n => n.ParentId == folder.Id).Take(take).Select(ToSummary)]);
    }

    public Task<IReadOnlyList<DriveItemSummary>> SearchAsync(
        string driveId, string query, int top = 25, CancellationToken ct = default)
    {
        var nodes = Drive(driveId);
        if (string.IsNullOrWhiteSpace(query))
        {
            throw GraphServiceException.InvalidRequest(
                "Query must not be empty.", "pass a filename or keyword");
        }

        LastSearch = query;
        int take = Math.Clamp(top, 1, 50);
        return Task.FromResult<IReadOnlyList<DriveItemSummary>>(
            [.. nodes.Values
                .Where(n => n.Name.Contains(query, StringComparison.OrdinalIgnoreCase))
                .Take(take).Select(ToSummary)]);
    }

    public Task<FileContentDto> DownloadAsync(
        string driveId, string itemRef, int maxBytes = 786432, CancellationToken ct = default)
    {
        var node = Get(driveId, itemRef);
        if (node.IsFolder)
        {
            throw GraphServiceException.InvalidRequest(
                $"Drive item '{node.Name}' is a folder.",
                "download files only; use sharepoint_list_children to browse folders");
        }

        int cap = Math.Clamp(maxBytes, 1, 2097152);
        if (node.Content.Length > cap)
        {
            throw GraphServiceException.AttachmentTooLarge(node.Name, node.Content.Length, cap);
        }

        if (IsText(node.MimeType))
        {
            string text = System.Text.Encoding.UTF8.GetString(node.Content);
            bool truncated = text.Length > 20000;
            return Task.FromResult(new FileContentDto(
                node.Id, node.Name, node.MimeType, node.Content.Length, "text",
                truncated ? text[..20000] + "…[truncated]" : text, null, truncated));
        }

        return Task.FromResult(new FileContentDto(
            node.Id, node.Name, node.MimeType, node.Content.Length, "base64",
            null, Convert.ToBase64String(node.Content), false));
    }

    public Task<FileContentDto> DownloadToFileAsync(
        string driveId, string itemRef, string localPath, bool overwrite = false, CancellationToken ct = default)
    {
        string dest = PathResolver.RequireAbsolutePath(localPath);
        if (_fileSystem.File.Exists(dest) && !overwrite)
        {
            throw GraphServiceException.InvalidRequest(
                $"Local file '{dest}' already exists.",
                "pass overwrite:true to replace it, or choose another localPath");
        }

        var node = Get(driveId, itemRef);
        if (node.IsFolder)
        {
            throw GraphServiceException.InvalidRequest(
                $"Drive item '{node.Name}' is a folder.",
                "download files only; use sharepoint_list_children to browse folders");
        }

        string? directory = _fileSystem.Path.GetDirectoryName(dest);
        if (directory is not null)
        {
            _fileSystem.Directory.CreateDirectory(directory);
        }

        _fileSystem.File.WriteAllBytes(dest, node.Content);
        return Task.FromResult(new FileContentDto(
            node.Id, node.Name, node.MimeType, node.Content.Length, "file",
            null, null, false, dest));
    }

    public Task<DriveItemSummary> CreateFolderAsync(
        string driveId, string name, string parentRef = "root", CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(name) || name.IndexOfAny(['/', '\\']) >= 0)
        {
            throw GraphServiceException.InvalidRequest(
                "Folder name must be non-empty and contain no slashes.",
                "pass a plain name, e.g. \"Belege\"");
        }

        var parent = Get(driveId, parentRef);
        var folder = new Node
        {
            Id = $"f-new-{++_counter}",
            Name = name.Trim(),
            IsFolder = true,
            ParentId = parent.Id
        };
        Add(driveId.Trim(), folder);
        return Task.FromResult(ToSummary(folder));
    }

    public Task<DriveItemSummary> UploadAsync(
        string driveId,
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

        byte[] bytes;
        if (localPath is not null)
        {
            if (contentText is not null || contentBase64 is not null)
            {
                throw GraphServiceException.InvalidRequest(
                    "Provide either localPath or inline content, not both.",
                    "use localPath for files on disk, contentText for small snippets");
            }

            string full = PathResolver.RequireAbsolutePath(localPath);
            if (!_fileSystem.File.Exists(full))
            {
                throw GraphServiceException.InvalidRequest(
                    $"Local file '{full}' does not exist.",
                    "pass the absolute path of an existing file on the host");
            }

            bytes = _fileSystem.File.ReadAllBytes(full);
            if (bytes.Length > 104_857_600)
            {
                throw GraphServiceException.InvalidRequest(
                    $"Local file is {bytes.Length} bytes, above the 104857600 byte tool limit.",
                    "upload it via the SharePoint UI (API sessions support up to 250 GB)");
            }
        }
        else
        {
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

            if (bytes.Length > 4_194_304)
            {
                throw GraphServiceException.InvalidRequest(
                    $"Inline content is {bytes.Length} bytes, above the 4194304 byte simple-upload limit.",
                    "write the content to a file and pass its absolute localPath (resumable upload)");
            }
        }

        var parent = Get(driveId, parentRef ?? "root");
        var file = new Node
        {
            Id = $"f-up-{++_counter}",
            Name = fileName.Trim(),
            ParentId = parent.Id,
            MimeType = "text/plain",
            Content = bytes
        };
        Add(driveId.Trim(), file);
        return Task.FromResult(ToSummary(file));
    }

    public Task<DriveItemSummary> MoveAsync(
        string driveId,
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

        var node = Get(driveId, itemId);
        if (newParentRef is not null)
        {
            node.ParentId = Get(driveId, newParentRef).Id;
        }

        if (newName is not null)
        {
            node.Name = newName;
        }

        return Task.FromResult(ToSummary(node));
    }

    // Mirrors GraphSharePointService text rules so the fake behaves alike.
    private static bool IsText(string? mimeType)
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
