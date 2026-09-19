using System.Text.Json.Serialization;

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

public sealed record FolderInfo
{
    public string Id { get; init; }
    public string DisplayName { get; init; }
    public int TotalCount { get; init; }
    public int UnreadCount { get; init; }
    public string? ParentId { get; init; }
    public string? Path { get; init; }

    public FolderInfo(
        string id,
        string displayName,
        int totalCount,
        int unreadCount)
    {
        Id = id;
        DisplayName = displayName;
        TotalCount = totalCount;
        UnreadCount = unreadCount;
    }

    [JsonConstructor]
    public FolderInfo(
        string id,
        string displayName,
        int totalCount,
        int unreadCount,
        string? parentId,
        string? path)
        : this(id, displayName, totalCount, unreadCount)
    {
        ParentId = parentId;
        Path = path;
    }

    public void Deconstruct(
        out string id,
        out string displayName,
        out int totalCount,
        out int unreadCount)
    {
        id = Id;
        displayName = DisplayName;
        totalCount = TotalCount;
        unreadCount = UnreadCount;
    }
}

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

/// <summary>Attachment content. Encoding is text, base64, reference (a
/// OneDrive or other storage link, not downloaded; Graph v1.0 may not expose
/// the URL) or nested (message/event attachment, not downloaded).</summary>
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
