using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using MicrosoftMcp.Common;

namespace MicrosoftMcp.Outlook;

public sealed class GraphMailService(
    GraphServiceClient client,
    IOptions<GraphAuthOptions> options) : IGraphMailService
{
    private static readonly string[] SummarySelect =
        ["id", "subject", "from", "toRecipients", "receivedDateTime", "isRead",
         "hasAttachments", "categories", "importance", "bodyPreview", "parentFolderId", "webLink"];

    private readonly GraphAuthOptions _options = options.Value;
    private bool IsMe => string.Equals(_options.UserIdOrUpn, "me", StringComparison.OrdinalIgnoreCase);

    public async Task<IReadOnlyList<EmailSummary>> SearchAsync(EmailQuery query, CancellationToken ct = default)
    {
        int top = Math.Clamp(query.Top, 1, 50);
        string? folderId = string.IsNullOrWhiteSpace(query.Folder)
            ? null
            : await ResolveFolderIdAsync(query.Folder!, ct).ConfigureAwait(false);

        List<Message> items = folderId is null
            ? await SearchAcrossMailboxAsync(query, top, ct).ConfigureAwait(false)
            : await SearchInFolderAsync(folderId, query, top, ct).ConfigureAwait(false);

        return [.. items.Select(EmailMapper.MapSummary)];
    }

    public async Task<EmailDetail> GetAsync(string messageId, CancellationToken ct = default)
    {
        RequireId(messageId);
        Message? msg = IsMe
            ? await client.Me.Messages[messageId].GetAsync(c =>
                c.QueryParameters.Select = [.. SummarySelect, "body"], ct).ConfigureAwait(false)
            : await client.Users[_options.UserIdOrUpn].Messages[messageId].GetAsync(c =>
                c.QueryParameters.Select = [.. SummarySelect, "body"], ct).ConfigureAwait(false);

        return msg is null
            ? throw MailServiceException.MessageNotFound(messageId, "outlook_read_email")
            : EmailMapper.MapDetail(msg);
    }

    public async Task<IReadOnlyList<FolderInfo>> ListFoldersAsync(CancellationToken ct = default)
    {
        if (IsMe)
        {
            var page = await client.Me.MailFolders.GetAsync(c => c.QueryParameters.Top = 100, ct).ConfigureAwait(false);
            return [.. (page?.Value ?? []).Select(MapFolder)];
        }

        var userPage = await client.Users[_options.UserIdOrUpn].MailFolders
            .GetAsync(c => c.QueryParameters.Top = 100, ct).ConfigureAwait(false);
        return [.. (userPage?.Value ?? []).Select(MapFolder)];
    }

    public async Task<FolderInfo> CreateFolderAsync(
        string displayName, string? parentFolderId = null, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(displayName);
        var folder = new MailFolder { DisplayName = displayName.Trim(), ParentFolderId = parentFolderId };

        MailFolder? created = IsMe
            ? await client.Me.MailFolders.PostAsync(folder, cancellationToken: ct).ConfigureAwait(false)
            : await client.Users[_options.UserIdOrUpn].MailFolders
                .PostAsync(folder, cancellationToken: ct).ConfigureAwait(false);

        return created is null
            ? throw MailServiceException.GraphError(0, null, "Folder creation returned no result.")
            : MapFolder(created);
    }

    public async Task<EmailSummary> MoveAsync(string messageId, string destination, CancellationToken ct = default)
    {
        RequireId(messageId);
        string destId = await ResolveFolderIdAsync(destination, ct).ConfigureAwait(false);

        if (IsMe)
        {
            var body = new Microsoft.Graph.Me.Messages.Item.Move.MovePostRequestBody { DestinationId = destId };
            var moved = await client.Me.Messages[messageId].Move.PostAsync(body, cancellationToken: ct).ConfigureAwait(false);
            return moved is null
                ? throw MailServiceException.MessageNotFound(messageId, "outlook_move_email")
                : EmailMapper.MapSummary(moved);
        }

        var userBody = new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody { DestinationId = destId };
        var userMoved = await client.Users[_options.UserIdOrUpn].Messages[messageId].Move
            .PostAsync(userBody, cancellationToken: ct).ConfigureAwait(false);
        return userMoved is null
            ? throw MailServiceException.MessageNotFound(messageId, "outlook_move_email")
            : EmailMapper.MapSummary(userMoved);
    }

    public Task<EmailSummary> ArchiveAsync(string messageId, CancellationToken ct = default) =>
        MoveAsync(messageId, "archive", ct);

    public async Task DeleteAsync(string messageId, CancellationToken ct = default)
    {
        // Soft-delete: move to DeletedItems (reversible). No hard delete offered.
        await MoveAsync(messageId, "deleteditems", ct).ConfigureAwait(false);
    }

    public async Task<EmailDetail> CreateDraftAsync(
        IReadOnlyList<string> to,
        string subject,
        string body,
        bool isHtml = false,
        CancellationToken ct = default)
    {
        RequireRecipients(to);
        RequireBody(body);

        var msg = new Message
        {
            Subject = subject,
            Body = new ItemBody
            {
                ContentType = isHtml ? BodyType.Html : BodyType.Text,
                Content = body
            },
            ToRecipients = [.. to.Select(ToRecipient)]
        };

        Message? draft = IsMe
            ? await client.Me.Messages.PostAsync(msg, cancellationToken: ct).ConfigureAwait(false)
            : await client.Users[_options.UserIdOrUpn].Messages.PostAsync(msg, cancellationToken: ct).ConfigureAwait(false);

        return draft is null
            ? throw MailServiceException.GraphError(0, null, "Draft creation returned no result.")
            : EmailMapper.MapDetail(draft);
    }

    public async Task<EmailDetail> CreateReplyDraftAsync(
        string messageId, string comment, CancellationToken ct = default)
    {
        RequireId(messageId);
        RequireBody(comment);

        if (IsMe)
        {
            var body = new Microsoft.Graph.Me.Messages.Item.CreateReply.CreateReplyPostRequestBody
            {
                Comment = comment
            };
            var draft = await client.Me.Messages[messageId].CreateReply
                .PostAsync(body, cancellationToken: ct).ConfigureAwait(false);
            return draft is null
                ? throw MailServiceException.MessageNotFound(messageId, "outlook_create_reply_draft")
                : EmailMapper.MapDetail(draft);
        }

        var userBody = new Microsoft.Graph.Users.Item.Messages.Item.CreateReply.CreateReplyPostRequestBody
        {
            Comment = comment
        };
        var userDraft = await client.Users[_options.UserIdOrUpn].Messages[messageId].CreateReply
            .PostAsync(userBody, cancellationToken: ct).ConfigureAwait(false);
        return userDraft is null
            ? throw MailServiceException.MessageNotFound(messageId, "outlook_create_reply_draft")
            : EmailMapper.MapDetail(userDraft);
    }

    public async Task<EmailDetail> CreateForwardDraftAsync(
        string messageId, IReadOnlyList<string> to, string? comment = null, CancellationToken ct = default)
    {
        RequireId(messageId);
        RequireRecipients(to);

        if (IsMe)
        {
            var body = new Microsoft.Graph.Me.Messages.Item.CreateForward.CreateForwardPostRequestBody
            {
                ToRecipients = [.. to.Select(ToRecipient)],
                Comment = comment
            };
            var draft = await client.Me.Messages[messageId].CreateForward
                .PostAsync(body, cancellationToken: ct).ConfigureAwait(false);
            return draft is null
                ? throw MailServiceException.MessageNotFound(messageId, "outlook_create_forward_draft")
                : EmailMapper.MapDetail(draft);
        }

        var userBody = new Microsoft.Graph.Users.Item.Messages.Item.CreateForward.CreateForwardPostRequestBody
        {
            ToRecipients = [.. to.Select(ToRecipient)],
            Comment = comment
        };
        var userDraft = await client.Users[_options.UserIdOrUpn].Messages[messageId].CreateForward
            .PostAsync(userBody, cancellationToken: ct).ConfigureAwait(false);
        return userDraft is null
            ? throw MailServiceException.MessageNotFound(messageId, "outlook_create_forward_draft")
            : EmailMapper.MapDetail(userDraft);
    }

    public async Task<EmailDetail> UpdateDraftAsync(
        string messageId,
        string? subject = null,
        string? body = null,
        bool isHtml = false,
        IReadOnlyList<string>? to = null,
        CancellationToken ct = default)
    {
        RequireId(messageId);
        if (subject is null && body is null && to is null)
        {
            throw MailServiceException.InvalidRequest(
                "Nothing to update: provide at least one of subject, body or to.",
                "pass subject and/or body and/or to, then retry");
        }

        var patch = new Message();
        if (subject is not null)
        {
            patch.Subject = subject;
        }

        if (body is not null)
        {
            patch.Body = new ItemBody
            {
                ContentType = isHtml ? BodyType.Html : BodyType.Text,
                Content = body
            };
        }

        if (to is not null)
        {
            patch.ToRecipients = [.. to.Select(ToRecipient)];
        }

        if (IsMe)
        {
            var updated = await client.Me.Messages[messageId]
                .PatchAsync(patch, cancellationToken: ct).ConfigureAwait(false);
            return updated is null
                ? throw MailServiceException.MessageNotFound(messageId, "outlook_update_draft")
                : EmailMapper.MapDetail(updated);
        }

        var userUpdated = await client.Users[_options.UserIdOrUpn].Messages[messageId]
            .PatchAsync(patch, cancellationToken: ct).ConfigureAwait(false);
        return userUpdated is null
            ? throw MailServiceException.MessageNotFound(messageId, "outlook_update_draft")
            : EmailMapper.MapDetail(userUpdated);
    }

    public async Task<IReadOnlyList<AttachmentInfo>> ListAttachmentsAsync(
        string messageId, CancellationToken ct = default)
    {
        RequireId(messageId);
        if (IsMe)
        {
            var page = await client.Me.Messages[messageId].Attachments.GetAsync(c =>
            {
                c.QueryParameters.Top = 100;
                c.QueryParameters.Select = ["id", "name", "contentType", "size", "isInline"];
            }, ct).ConfigureAwait(false);
            return [.. (page?.Value ?? []).Select(EmailMapper.MapAttachment)];
        }

        var userPage = await client.Users[_options.UserIdOrUpn].Messages[messageId].Attachments.GetAsync(c =>
        {
            c.QueryParameters.Top = 100;
            c.QueryParameters.Select = ["id", "name", "contentType", "size", "isInline"];
        }, ct).ConfigureAwait(false);
        return [.. (userPage?.Value ?? []).Select(EmailMapper.MapAttachment)];
    }

    public async Task<AttachmentContent> ReadAttachmentAsync(
        string messageId, string attachmentId, int maxBytes = 786432, CancellationToken ct = default)
    {
        RequireId(messageId);
        if (string.IsNullOrWhiteSpace(attachmentId))
        {
            throw MailServiceException.MissingId("attachmentId");
        }

        int cap = Math.Clamp(maxBytes, 1, 2097152);

        Attachment? att = IsMe
            ? await client.Me.Messages[messageId].Attachments[attachmentId].GetAsync(c =>
                c.QueryParameters.Select = ["id", "name", "contentType", "size", "isInline", "contentBytes"],
                ct).ConfigureAwait(false)
            : await client.Users[_options.UserIdOrUpn].Messages[messageId].Attachments[attachmentId].GetAsync(c =>
                c.QueryParameters.Select = ["id", "name", "contentType", "size", "isInline", "contentBytes"],
                ct).ConfigureAwait(false);

        if (att is null)
        {
            throw MailServiceException.InvalidRequest(
                $"Attachment '{attachmentId}' was not found on message '{messageId}'.",
                "call outlook_list_attachments to get valid attachment ids");
        }

        return att switch
        {
            FileAttachment file => MapFileAttachment(file, cap),
            ItemAttachment item => new AttachmentContent(
                item.Id ?? string.Empty, item.Name ?? string.Empty, item.ContentType,
                item.Size ?? 0, "nested", null, null, null, false),
            ReferenceAttachment reference => new AttachmentContent(
                reference.Id ?? string.Empty, reference.Name ?? string.Empty, reference.ContentType,
                reference.Size ?? 0, "reference", null, null, ReferenceUrl(reference), false),
            _ => new AttachmentContent(
                att.Id ?? string.Empty, att.Name ?? string.Empty, att.ContentType,
                att.Size ?? 0, "unknown", null, null, null, false)
        };
    }

    private static AttachmentContent MapFileAttachment(FileAttachment file, int cap)
    {
        byte[] bytes = file.ContentBytes ?? [];
        if (bytes.Length > cap)
        {
            throw MailServiceException.AttachmentTooLarge(file.Name ?? "?", bytes.Length, cap);
        }

        if (IsTextContent(file.ContentType))
        {
            string text = System.Text.Encoding.UTF8.GetString(bytes);
            bool truncated = text.Length > 20000;
            return new AttachmentContent(
                file.Id ?? string.Empty, file.Name ?? string.Empty, file.ContentType,
                bytes.Length, "text",
                truncated ? EmailMapper.Truncate(text, 20000) : text,
                null, null, truncated);
        }

        return new AttachmentContent(
            file.Id ?? string.Empty, file.Name ?? string.Empty, file.ContentType,
            bytes.Length, "base64", null, Convert.ToBase64String(bytes), null, false);
    }

    private static string? ReferenceUrl(ReferenceAttachment reference) =>
        reference.AdditionalData.TryGetValue("sourceUrl", out var url) ? url?.ToString() : null;

    private static bool IsTextContent(string? contentType)
    {
        if (string.IsNullOrWhiteSpace(contentType))
        {
            return false;
        }

        string type = contentType.Split(';')[0].Trim().ToLowerInvariant();
        return type.StartsWith("text/", StringComparison.Ordinal)
            || type is "application/json" or "application/xml"
                or "application/javascript" or "application/csv"
            || type.EndsWith("+json", StringComparison.Ordinal)
            || type.EndsWith("+xml", StringComparison.Ordinal);
    }

    public async Task<IReadOnlyList<CategoryInfo>> ListCategoriesAsync(CancellationToken ct = default)
    {
        if (IsMe)
        {
            var page = await client.Me.Outlook.MasterCategories
                .GetAsync(c => c.QueryParameters.Top = 100, ct).ConfigureAwait(false);
            return [.. (page?.Value ?? []).Select(EmailMapper.MapCategory)];
        }

        var userPage = await client.Users[_options.UserIdOrUpn].Outlook.MasterCategories
            .GetAsync(c => c.QueryParameters.Top = 100, ct).ConfigureAwait(false);
        return [.. (userPage?.Value ?? []).Select(EmailMapper.MapCategory)];
    }

    public async Task<EmailDetail> SetCategoriesAsync(
        string messageId, string[] add, string[] remove, CancellationToken ct = default)
    {
        RequireId(messageId);
        var current = await GetAsync(messageId, ct).ConfigureAwait(false);
        var patch = new Message
        {
            Categories = EmailMapper.MergeCategories(current.Categories, add, remove)
        };

        if (IsMe)
        {
            var updated = await client.Me.Messages[messageId].PatchAsync(patch, cancellationToken: ct).ConfigureAwait(false);
            return updated is null
                ? throw MailServiceException.MessageNotFound(messageId, "outlook_set_categories")
                : EmailMapper.MapDetail(updated);
        }

        var userUpdated = await client.Users[_options.UserIdOrUpn].Messages[messageId]
            .PatchAsync(patch, cancellationToken: ct).ConfigureAwait(false);
        return userUpdated is null
            ? throw MailServiceException.MessageNotFound(messageId, "outlook_set_categories")
            : EmailMapper.MapDetail(userUpdated);
    }

    public async Task MarkReadAsync(string messageId, bool isRead, CancellationToken ct = default)
    {
        RequireId(messageId);
        var patch = new Message { IsRead = isRead };
        if (IsMe)
        {
            await client.Me.Messages[messageId].PatchAsync(patch, cancellationToken: ct).ConfigureAwait(false);
            return;
        }

        await client.Users[_options.UserIdOrUpn].Messages[messageId]
            .PatchAsync(patch, cancellationToken: ct).ConfigureAwait(false);
    }

    public async Task SetImportanceAsync(string messageId, string importance, CancellationToken ct = default)
    {
        RequireId(messageId);
        if (!Enum.TryParse<Importance>(importance, ignoreCase: true, out var level))
        {
            throw MailServiceException.InvalidRequest(
                $"Invalid importance '{importance}'.",
                "use low, normal or high");
        }

        var patch = new Message { Importance = level };
        if (IsMe)
        {
            await client.Me.Messages[messageId].PatchAsync(patch, cancellationToken: ct).ConfigureAwait(false);
            return;
        }

        await client.Users[_options.UserIdOrUpn].Messages[messageId]
            .PatchAsync(patch, cancellationToken: ct).ConfigureAwait(false);
    }

    private async Task<List<Message>> SearchAcrossMailboxAsync(
        EmailQuery query, int top, CancellationToken ct)
    {
        if (!string.IsNullOrWhiteSpace(query.Query))
        {
            if (IsMe)
            {
                var page = await client.Me.Messages.GetAsync(c =>
                {
                    c.QueryParameters.Top = top;
                    c.QueryParameters.Search = $"\"{query.Query}\"";
                    c.QueryParameters.Select = SummarySelect;
                }, ct).ConfigureAwait(false);
                return [.. page?.Value ?? []];
            }

            var userPage = await client.Users[_options.UserIdOrUpn].Messages.GetAsync(c =>
            {
                c.QueryParameters.Top = top;
                c.QueryParameters.Search = $"\"{query.Query}\"";
                c.QueryParameters.Select = SummarySelect;
            }, ct).ConfigureAwait(false);
            return [.. userPage?.Value ?? []];
        }

        string? filter = string.IsNullOrWhiteSpace(query.From)
            ? null
            : $"from/emailAddress/address eq '{query.From!.Replace("'", "''")}'";
        if (IsMe)
        {
            var page = await client.Me.Messages.GetAsync(c =>
            {
                c.QueryParameters.Top = top;
                c.QueryParameters.Filter = filter;
                c.QueryParameters.Orderby = ["receivedDateTime desc"];
                c.QueryParameters.Select = SummarySelect;
            }, ct).ConfigureAwait(false);
            return [.. page?.Value ?? []];
        }

        var filtered = await client.Users[_options.UserIdOrUpn].Messages.GetAsync(c =>
        {
            c.QueryParameters.Top = top;
            c.QueryParameters.Filter = filter;
            c.QueryParameters.Orderby = ["receivedDateTime desc"];
            c.QueryParameters.Select = SummarySelect;
        }, ct).ConfigureAwait(false);
        return [.. filtered?.Value ?? []];
    }

    private async Task<List<Message>> SearchInFolderAsync(
        string folderId, EmailQuery query, int top, CancellationToken ct)
    {
        if (!string.IsNullOrWhiteSpace(query.Query))
        {
            if (IsMe)
            {
                var page = await client.Me.MailFolders[folderId].Messages.GetAsync(c =>
                {
                    c.QueryParameters.Top = top;
                    c.QueryParameters.Search = $"\"{query.Query}\"";
                    c.QueryParameters.Select = SummarySelect;
                }, ct).ConfigureAwait(false);
                return [.. page?.Value ?? []];
            }

            var userPage = await client.Users[_options.UserIdOrUpn].MailFolders[folderId].Messages.GetAsync(c =>
            {
                c.QueryParameters.Top = top;
                c.QueryParameters.Search = $"\"{query.Query}\"";
                c.QueryParameters.Select = SummarySelect;
            }, ct).ConfigureAwait(false);
            return [.. userPage?.Value ?? []];
        }

        string? filter = string.IsNullOrWhiteSpace(query.From)
            ? null
            : $"from/emailAddress/address eq '{query.From!.Replace("'", "''")}'";
        if (IsMe)
        {
            var page = await client.Me.MailFolders[folderId].Messages.GetAsync(c =>
            {
                c.QueryParameters.Top = top;
                c.QueryParameters.Filter = filter;
                c.QueryParameters.Orderby = ["receivedDateTime desc"];
                c.QueryParameters.Select = SummarySelect;
            }, ct).ConfigureAwait(false);
            return [.. page?.Value ?? []];
        }

        var filtered = await client.Users[_options.UserIdOrUpn].MailFolders[folderId].Messages.GetAsync(c =>
        {
            c.QueryParameters.Top = top;
            c.QueryParameters.Filter = filter;
            c.QueryParameters.Orderby = ["receivedDateTime desc"];
            c.QueryParameters.Select = SummarySelect;
        }, ct).ConfigureAwait(false);
        return [.. filtered?.Value ?? []];
    }

    private async Task<string> ResolveFolderIdAsync(string destination, CancellationToken ct)
    {
        if (FolderResolver.IsWellKnown(destination))
        {
            return destination.Trim().ToLowerInvariant();
        }

        var folders = await ListFoldersAsync(ct).ConfigureAwait(false);
        return FolderResolver.Resolve(destination, folders);
    }

    private static void RequireId(string messageId)
    {
        if (string.IsNullOrWhiteSpace(messageId))
        {
            throw MailServiceException.MissingId();
        }
    }

    private static void RequireRecipients(IReadOnlyList<string> to)
    {
        if (to.Count == 0 || to.Any(string.IsNullOrWhiteSpace))
        {
            throw MailServiceException.InvalidRequest(
                "At least one valid recipient address is required.",
                "pass non-empty 'to' addresses, e.g. [\"a@example.com\"]");
        }
    }

    private static void RequireBody(string? body)
    {
        if (string.IsNullOrWhiteSpace(body))
        {
            throw MailServiceException.InvalidRequest(
                "Body must not be empty.",
                "provide the mail text (or html with isHtml=true)");
        }
    }

    private static FolderInfo MapFolder(MailFolder f) => new(
        f.Id ?? string.Empty,
        f.DisplayName ?? string.Empty,
        f.TotalItemCount ?? 0,
        f.UnreadItemCount ?? 0);

    private static Recipient ToRecipient(string address) => new()
    {
        EmailAddress = new EmailAddress { Address = address }
    };
}
