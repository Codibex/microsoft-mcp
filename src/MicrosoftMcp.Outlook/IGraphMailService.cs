namespace MicrosoftMcp.Outlook;

public interface IGraphMailService
{
    Task<IReadOnlyList<EmailSummary>> SearchAsync(EmailQuery query, CancellationToken ct = default);
    Task<EmailDetail> GetAsync(string messageId, CancellationToken ct = default);
    Task<IReadOnlyList<FolderInfo>> ListFoldersAsync(CancellationToken ct = default);
    Task<FolderInfo> CreateFolderAsync(string displayName, string? parentFolderId = null, CancellationToken ct = default);
    Task<EmailSummary> MoveAsync(string messageId, string destination, CancellationToken ct = default);
    Task<EmailSummary> ArchiveAsync(string messageId, CancellationToken ct = default);
    Task DeleteAsync(string messageId, CancellationToken ct = default);
    Task<EmailDetail> CreateDraftAsync(
        IReadOnlyList<string> to,
        string subject,
        string body,
        bool isHtml = false,
        CancellationToken ct = default);
    Task<EmailDetail> CreateReplyDraftAsync(
        string messageId,
        string comment,
        CancellationToken ct = default);
    Task<EmailDetail> CreateForwardDraftAsync(
        string messageId,
        IReadOnlyList<string> to,
        string? comment = null,
        CancellationToken ct = default);
    Task<EmailDetail> UpdateDraftAsync(
        string messageId,
        string? subject = null,
        string? body = null,
        bool isHtml = false,
        IReadOnlyList<string>? to = null,
        CancellationToken ct = default);
    Task<IReadOnlyList<AttachmentInfo>> ListAttachmentsAsync(string messageId, CancellationToken ct = default);
    Task<AttachmentContent> ReadAttachmentAsync(
        string messageId,
        string attachmentId,
        int maxBytes = 786432,
        CancellationToken ct = default);
    Task<IReadOnlyList<CategoryInfo>> ListCategoriesAsync(CancellationToken ct = default);
    Task<EmailDetail> SetCategoriesAsync(string messageId, string[] add, string[] remove, CancellationToken ct = default);
    Task MarkReadAsync(string messageId, bool isRead, CancellationToken ct = default);
    Task SetImportanceAsync(string messageId, string importance, CancellationToken ct = default);
}
