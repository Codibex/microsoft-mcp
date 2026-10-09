namespace MicrosoftMcp.SharePoint;

public sealed record SharePointSiteInfo(
    string Id,
    string Name,
    string? DisplayName,
    string? WebUrl);

public sealed record SharePointDriveInfo(
    string Id,
    string Name,
    string? DriveType,
    string? WebUrl,
    long TotalBytes,
    long UsedBytes,
    long RemainingBytes);

public sealed record DriveItemSummary(
    string Id,
    string Name,
    bool IsFolder,
    long Size,
    string? MimeType,
    string? ParentId,
    string? WebUrl,
    DateTimeOffset? LastModified,
    int ChildCount);

/// <summary>File content. Encoding is text, base64 or file (saved to
/// <see cref="LocalPath"/> on the host; bytes bypass the model context).</summary>
public sealed record FileContentDto(
    string Id,
    string Name,
    string? MimeType,
    long Size,
    string Encoding,
    string? Text,
    string? DataBase64,
    bool Truncated,
    string? LocalPath = null);
