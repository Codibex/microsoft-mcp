namespace MicrosoftMcp.Outlook;

public sealed record EmailAddressDto(string Name, string Address);

public sealed record EmailSummary(
    string Id,
    string Subject,
    EmailAddressDto? From,
    IReadOnlyList<EmailAddressDto> ToRecipients,
    DateTimeOffset? Received,
    bool IsRead,
    bool HasAttachments,
    IReadOnlyList<string> Categories,
    string? Importance,
    string? Preview);

public sealed record EmailDetail(
    string Id,
    string Subject,
    EmailAddressDto? From,
    IReadOnlyList<EmailAddressDto> ToRecipients,
    DateTimeOffset? Received,
    bool IsRead,
    IReadOnlyList<string> Categories,
    string? BodyPreview,
    string? Body,
    string? WebLink);

public sealed record FolderInfo(
    string Id,
    string DisplayName,
    int TotalCount,
    int UnreadCount,
    string? ParentId = null,
    string? Path = null);

public sealed record AttachmentInfo(
    string Id,
    string Name,
    string? ContentType,
    int Size,
    bool IsInline);

public sealed record CategoryInfo(
    string Id,
    string DisplayName,
    string? Color);

/// <summary>Attachment content. Encoding is text, base64, reference (OneDrive
/// link, not downloaded) or nested (message/event attachment, not downloaded).</summary>
public sealed record AttachmentContent(
    string Id,
    string Name,
    string? ContentType,
    int Size,
    string Encoding,
    string? Text,
    string? DataBase64,
    string? SourceUrl,
    bool Truncated);

public sealed record EmailQuery(
    string? Query,
    string? Folder,
    string? From,
    int Top = 25);
