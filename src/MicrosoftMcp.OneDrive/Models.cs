namespace MicrosoftMcp.OneDrive;

public sealed record DriveInfoDto(
    string Id,
    string Name,
    string? DriveType,
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

/// <summary>File content. Encoding is text or base64 (see read limits).</summary>
public sealed record FileContentDto(
    string Id,
    string Name,
    string? MimeType,
    long Size,
    string Encoding,
    string? Text,
    string? DataBase64,
    bool Truncated);
