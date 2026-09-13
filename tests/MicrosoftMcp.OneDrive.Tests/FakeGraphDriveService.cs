namespace MicrosoftMcp.OneDrive.Tests;

using MicrosoftMcp.Common;

/// <summary>In-memory fake of <see cref="IGraphDriveService"/>. No Graph, no network.</summary>
internal sealed class FakeGraphDriveService : IGraphDriveService
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

    private readonly Dictionary<string, Node> _nodes = new(StringComparer.OrdinalIgnoreCase);
    private int _counter;

    public FakeGraphDriveService()
    {
        Add(new Node { Id = "root", Name = "root", IsFolder = true });
        Add(new Node { Id = "f-docs", Name = "Dokumente", IsFolder = true, ParentId = "root" });
        Add(new Node
        {
            Id = "f-note", Name = "notiz.txt", ParentId = "root",
            MimeType = "text/plain", Content = "Hallo Welt\nZeile zwei"u8.ToArray()
        });
        Add(new Node
        {
            Id = "f-img", Name = "bild.png", ParentId = "f-docs",
            MimeType = "image/png", Content = new byte[64]
        });
        Add(new Node
        {
            Id = "f-big", Name = "gross.bin", ParentId = "root",
            MimeType = "application/octet-stream", Content = new byte[3_000_000]
        });
    }

    public string? LastSearch { get; private set; }

    private void Add(Node n) => _nodes[n.Id] = n;

    private Node Get(string itemRef)
    {
        string reference = PathResolver.RequireRef(itemRef);
        if (PathResolver.IsRoot(reference))
        {
            return _nodes["root"];
        }

        if (PathResolver.IsPath(reference))
        {
            string[] parts = PathResolver.NormalizePath(reference).Split('/');
            Node current = _nodes["root"];
            foreach (string part in parts)
            {
                current = _nodes.Values.FirstOrDefault(n =>
                        string.Equals(n.ParentId, current.Id, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(n.Name, part, StringComparison.OrdinalIgnoreCase))
                    ?? throw MailServiceException.DriveItemNotFound(itemRef, "fake");
            }

            return current;
        }

        return _nodes.TryGetValue(reference, out var node)
            ? node
            : throw MailServiceException.DriveItemNotFound(itemRef, "fake");
    }

    private static DriveItemSummary ToSummary(Node n) => new(
        n.Id, n.Name, n.IsFolder, n.Content.Length,
        n.MimeType, n.ParentId, null, DateTimeOffset.UtcNow,
        n.IsFolder ? 1 : 0);

    public Task<DriveInfoDto> GetDriveAsync(CancellationToken ct = default) =>
        Task.FromResult(new DriveInfoDto("d1", "OneDrive", "personal", 1000, 100, 900));

    public Task<IReadOnlyList<DriveInfoDto>> ListDrivesAsync(CancellationToken ct = default) =>
        Task.FromResult<IReadOnlyList<DriveInfoDto>>(
            [new DriveInfoDto("d1", "OneDrive", "personal", 1000, 100, 900)]);

    public Task<DriveItemSummary> GetItemAsync(string itemRef, CancellationToken ct = default) =>
        Task.FromResult(ToSummary(Get(itemRef)));

    public Task<IReadOnlyList<DriveItemSummary>> ListChildrenAsync(
        string folderRef = "root", int top = 50, CancellationToken ct = default)
    {
        var folder = Get(folderRef);
        if (!folder.IsFolder)
        {
            throw MailServiceException.InvalidRequest(
                $"Drive item '{folder.Name}' is a folder expected.",
                "pass a folder id or /path");
        }

        int take = Math.Clamp(top, 1, 200);
        return Task.FromResult<IReadOnlyList<DriveItemSummary>>(
            [.. _nodes.Values.Where(n => n.ParentId == folder.Id).Take(take).Select(ToSummary)]);
    }

    public Task<IReadOnlyList<DriveItemSummary>> SearchAsync(
        string query, int top = 25, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            throw MailServiceException.InvalidRequest(
                "Query must not be empty.", "pass a filename or keyword");
        }

        LastSearch = query;
        int take = Math.Clamp(top, 1, 50);
        return Task.FromResult<IReadOnlyList<DriveItemSummary>>(
            [.. _nodes.Values
                .Where(n => n.Name.Contains(query, StringComparison.OrdinalIgnoreCase))
                .Take(take).Select(ToSummary)]);
    }

    public Task<FileContentDto> DownloadAsync(
        string itemRef, int maxBytes = 786432, CancellationToken ct = default)
    {
        var node = Get(itemRef);
        if (node.IsFolder)
        {
            throw MailServiceException.InvalidRequest(
                $"Drive item '{node.Name}' is a folder.",
                "download files only; use onedrive_list_children to browse folders");
        }

        int cap = Math.Clamp(maxBytes, 1, 2097152);
        if (node.Content.Length > cap)
        {
            throw MailServiceException.AttachmentTooLarge(node.Name, node.Content.Length, cap);
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

    public Task<DriveItemSummary> CreateFolderAsync(
        string name, string parentRef = "root", CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(name) || name.IndexOfAny(['/', '\\']) >= 0)
        {
            throw MailServiceException.InvalidRequest(
                "Folder name must be non-empty and contain no slashes.",
                "pass a plain name, e.g. \"Belege\"");
        }

        var parent = Get(parentRef);
        var folder = new Node
        {
            Id = $"f-new-{++_counter}",
            Name = name.Trim(),
            IsFolder = true,
            ParentId = parent.Id
        };
        Add(folder);
        return Task.FromResult(ToSummary(folder));
    }

    public Task<DriveItemSummary> UploadAsync(
        string fileName,
        string? parentRef = null,
        string? contentText = null,
        string? contentBase64 = null,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(fileName) || fileName.IndexOfAny(['/', '\\']) >= 0)
        {
            throw MailServiceException.InvalidRequest(
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
                    : throw MailServiceException.InvalidRequest(
                        "Either contentText or contentBase64 is required.",
                        "pass text directly or base64 for binary files");
        }
        catch (FormatException)
        {
            throw MailServiceException.InvalidRequest(
                "contentBase64 is not valid base64.",
                "pass correctly padded base64, or use contentText for text files");
        }

        if (bytes.Length > 4_194_304)
        {
            throw MailServiceException.InvalidRequest(
                $"File is {bytes.Length} bytes, above the 4194304 byte simple-upload limit.",
                "split the file or upload it via OneDrive UI");
        }

        var parent = Get(parentRef ?? "root");
        var file = new Node
        {
            Id = $"f-up-{++_counter}",
            Name = fileName.Trim(),
            ParentId = parent.Id,
            MimeType = "text/plain",
            Content = bytes
        };
        Add(file);
        return Task.FromResult(ToSummary(file));
    }

    public Task<DriveItemSummary> MoveAsync(
        string itemId,
        string? newParentRef = null,
        string? newName = null,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(itemId))
        {
            throw MailServiceException.MissingRef("itemId");
        }

        if (newParentRef is null && newName is null)
        {
            throw MailServiceException.InvalidRequest(
                "Nothing to do: provide newParentRef and/or newName.",
                "pass a destination folder and/or a new file name");
        }

        var node = Get(itemId);
        if (newParentRef is not null)
        {
            node.ParentId = Get(newParentRef).Id;
        }

        if (newName is not null)
        {
            node.Name = newName;
        }

        return Task.FromResult(ToSummary(node));
    }

    // Mirrors GraphDriveService text rules so the fake behaves alike.
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
