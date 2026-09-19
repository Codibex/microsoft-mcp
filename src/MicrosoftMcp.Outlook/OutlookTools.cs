using System.ComponentModel;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using MicrosoftMcp.Common;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace MicrosoftMcp.Outlook;

public static class OutlookServiceRegistration
{
    /// <param name="policy">Admin-owned Outlook policy from
    /// <see cref="MessagingPolicySetup.Initialize"/> (policy.json only, never env).
    /// Null = unrestricted defaults (no policy.json found).</param>
    public static IServiceCollection AddOutlook(this IServiceCollection services, OutlookPolicyOptions? policy = null)
    {
        services.AddSingleton<IGraphMailService, GraphMailService>();
        services.AddSingleton<IOptions<OutlookPolicyOptions>>(Options.Create(policy ?? new OutlookPolicyOptions()));
        return services;
    }
}

[McpServerToolType]
public sealed class OutlookTools(IGraphMailService mail, ILogger<OutlookTools> log)
{
    /// <summary>Single funnel: success becomes JSON text, any failure becomes an
    /// isError result with "[code] ... Next: ..." so the agent can act on it.
    /// Full detail (incl. stack) goes to the server log (stderr) only, never to
    /// the client. Cancellations still propagate.</summary>
    private async Task<CallToolResult> InvokeAsync<T>(string operation, Func<Task<T>> call)
    {
        try
        {
            return ToolResult.Ok(await call().ConfigureAwait(false));
        }
        catch (GraphServiceException ex)
        {
            log.LogWarning(ex, "Tool {Operation} failed with {Code}", operation, ex.Code);
            return ToolResult.Fail(ex);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            var mapped = GraphErrorMapper.ToGraphServiceException(ex, operation);
            log.LogError(ex, "Tool {Operation} failed unexpectedly ({Code})", operation, mapped.Code);
            return ToolResult.Fail(mapped);
        }
    }

    [McpServerTool, Description("Search mails via KQL ($search) or list recent. query e.g. 'subject:Rechnung from:boss'. folder optionally restricts to a folder (well-known name, id or displayName). top max 50. Failures return isError with [code] and a Next-hint.")]
    public Task<CallToolResult> outlook_search_emails(
        [Description("KQL query, empty lists recent")] string? query = null,
        [Description("Sender address filter")] string? from = null,
        [Description("Folder: well-known name, id or displayName")] string? folder = null,
        [Description("Max results 1-50")] int top = 25,
        CancellationToken ct = default) =>
        InvokeAsync("outlook_search_emails", () => mail.SearchAsync(new EmailQuery(query, folder, from, top), ct));

    [McpServerTool, Description("Read a full mail by Graph id. Body truncated to 8000 chars. Failures return isError with [code] and a Next-hint.")]
    public Task<CallToolResult> outlook_read_email(
        [Description("Graph message id")] string messageId,
        CancellationToken ct = default) =>
        InvokeAsync("outlook_read_email", () => mail.GetAsync(messageId, ct));

    [McpServerTool, Description("List mail folders with ids and counts. Needed to pick a move destination.")]
    public Task<CallToolResult> outlook_list_folders(CancellationToken ct = default) =>
        InvokeAsync("outlook_list_folders", () => mail.ListFoldersAsync(ct));

    [McpServerTool, Description("Create a mail folder, optionally under a parent folder. Returns the new folder with id.")]
    public Task<CallToolResult> outlook_create_folder(
        [Description("Display name of the new folder")] string displayName,
        [Description("Parent folder id (optional, omit for top level)")] string? parentFolderId = null,
        CancellationToken ct = default) =>
        InvokeAsync("outlook_create_folder", () => mail.CreateFolderAsync(displayName, parentFolderId, ct));

    [McpServerTool, Description("Move a mail to another folder. Destination: well-known name (inbox, archive, deleteditems, drafts), folder id or displayName. Reversible. Failures return isError with [code] and a Next-hint.")]
    public Task<CallToolResult> outlook_move_email(
        [Description("Graph message id")] string messageId,
        [Description("Destination folder")] string destination,
        CancellationToken ct = default) =>
        InvokeAsync("outlook_move_email", () => mail.MoveAsync(messageId, destination, ct));

    [McpServerTool, Description("Archive a mail (move to archive folder). Reversible triage action.")]
    public Task<CallToolResult> outlook_archive_email(
        [Description("Graph message id")] string messageId,
        CancellationToken ct = default) =>
        InvokeAsync("outlook_archive_email", () => mail.ArchiveAsync(messageId, ct));

    [McpServerTool, Description("Move a mail to trash (DeletedItems). Reversible, no hard delete.")]
    public Task<CallToolResult> outlook_delete_email(
        [Description("Graph message id")] string messageId,
        CancellationToken ct = default) =>
        InvokeAsync("outlook_delete_email", async () =>
        {
            await mail.DeleteAsync(messageId, ct).ConfigureAwait(false);
            return $"Message {messageId} moved to trash.";
        });

    [McpServerTool, Description("Create a draft (no send). Triage can prepare answers without risk.")]
    public Task<CallToolResult> outlook_create_draft(
        [Description("Recipient addresses")] string[] to,
        [Description("Subject")] string subject,
        [Description("Body text or html")] string body,
        [Description("True if body is html")] bool isHtml = false,
        CancellationToken ct = default) =>
        InvokeAsync("outlook_create_draft", () => mail.CreateDraftAsync(to, subject, body, isHtml, ct));

    [McpServerTool, Description("Create a reply draft for a mail (no send). The reply draft appears in Drafts for review.")]
    public Task<CallToolResult> outlook_create_reply_draft(
        [Description("Graph message id to reply to")] string messageId,
        [Description("Reply text (may contain html)")] string comment,
        CancellationToken ct = default) =>
        InvokeAsync("outlook_create_reply_draft", () => mail.CreateReplyDraftAsync(messageId, comment, ct));

    [McpServerTool, Description("Create a forward draft for a mail (no send). The forward draft appears in Drafts for review.")]
    public Task<CallToolResult> outlook_create_forward_draft(
        [Description("Graph message id to forward")] string messageId,
        [Description("Recipient addresses")] string[] to,
        [Description("Optional comment to prepend")] string? comment = null,
        CancellationToken ct = default) =>
        InvokeAsync("outlook_create_forward_draft", () => mail.CreateForwardDraftAsync(messageId, to, comment, ct));

    [McpServerTool, Description("Update a draft (subject, body, recipients). Only drafts can be updated.")]
    public Task<CallToolResult> outlook_update_draft(
        [Description("Graph message id of the draft")] string messageId,
        [Description("New subject (optional)")] string? subject = null,
        [Description("New body (optional)")] string? body = null,
        [Description("True if body is html")] bool isHtml = false,
        [Description("New recipients, replaces all (optional)")] string[]? to = null,
        CancellationToken ct = default) =>
        InvokeAsync("outlook_update_draft", () => mail.UpdateDraftAsync(messageId, subject, body, isHtml, to, ct));

    [McpServerTool, Description("List attachment metadata (id, name, type, size) of a mail. No content download.")]
    public Task<CallToolResult> outlook_list_attachments(
        [Description("Graph message id")] string messageId,
        CancellationToken ct = default) =>
        InvokeAsync("outlook_list_attachments", () => mail.ListAttachmentsAsync(messageId, ct));

    [McpServerTool, Description("Read an attachment. Text files come back decoded (truncated at 20000 chars), binary files as base64. Attachments above maxBytes are rejected with attachment-too-large. Nested messages and OneDrive links are reported, not downloaded.")]
    public Task<CallToolResult> outlook_read_attachment(
        [Description("Graph message id")] string messageId,
        [Description("Attachment id from outlook_list_attachments")] string attachmentId,
        [Description("Max bytes to download, up to 2097152")] int maxBytes = 786432,
        CancellationToken ct = default) =>
        InvokeAsync("outlook_read_attachment", () => mail.ReadAttachmentAsync(messageId, attachmentId, maxBytes, ct));

    [McpServerTool, Description("List available Outlook categories (master list) for labeling.")]
    public Task<CallToolResult> outlook_list_categories(CancellationToken ct = default) =>
        InvokeAsync("outlook_list_categories", () => mail.ListCategoriesAsync(ct));

    [McpServerTool, Description("Label a mail via Outlook categories (add/remove). Use outlook_list_categories to see available ones.")]
    public Task<CallToolResult> outlook_set_categories(
        [Description("Graph message id")] string messageId,
        [Description("Categories to add")] string[] add,
        [Description("Categories to remove")] string[] remove,
        CancellationToken ct = default) =>
        InvokeAsync("outlook_set_categories", () => mail.SetCategoriesAsync(messageId, add ?? [], remove ?? [], ct));

    [McpServerTool, Description("Mark a mail read/unread.")]
    public Task<CallToolResult> outlook_mark_read(
        [Description("Graph message id")] string messageId,
        [Description("True=read, false=unread")] bool isRead,
        CancellationToken ct = default) =>
        InvokeAsync("outlook_mark_read", async () =>
        {
            await mail.MarkReadAsync(messageId, isRead, ct).ConfigureAwait(false);
            return $"Message {messageId} marked {(isRead ? "read" : "unread")}.";
        });

    [McpServerTool, Description("Set mail importance (low, normal, high). Useful for triage prioritization.")]
    public Task<CallToolResult> outlook_set_importance(
        [Description("Graph message id")] string messageId,
        [Description("low, normal or high")] string importance,
        CancellationToken ct = default) =>
        InvokeAsync("outlook_set_importance", async () =>
        {
            await mail.SetImportanceAsync(messageId, importance, ct).ConfigureAwait(false);
            return $"Message {messageId} importance set to {importance}.";
        });
}
