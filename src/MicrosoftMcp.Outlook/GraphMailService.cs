using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using MicrosoftMcp.Common;

namespace MicrosoftMcp.Outlook;

public sealed class GraphMailService(
    GraphServiceClient client,
    IOptions<GraphAuthOptions> options,
    IOptions<OutlookPolicyOptions> policy) : IGraphMailService
{
    private static readonly string[] FolderSelect =
        ["id", "displayName", "parentFolderId", "totalItemCount", "unreadItemCount", "childFolderCount"];

    private static readonly string[] SummarySelect =
        ["id", "subject", "from", "toRecipients", "receivedDateTime", "isRead",
         "hasAttachments", "categories", "importance", "bodyPreview", "parentFolderId", "webLink"];

    private readonly GraphAuthOptions _options = options.Value;
    private readonly OutlookPolicyOptions _policy = policy.Value;
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
            ? throw GraphServiceException.MessageNotFound(messageId, "outlook_read_email")
            : EmailMapper.MapDetail(msg);
    }

    public Task<IReadOnlyList<FolderInfo>> ListFoldersAsync(CancellationToken ct = default) =>
        IsMe
            ? ListFoldersAsync(
                nextLink => GetMeRootPageAsync(nextLink, ct),
                (parentId, nextLink) => GetMeChildPageAsync(parentId, nextLink, ct))
            : ListFoldersAsync(
                nextLink => GetUserRootPageAsync(nextLink, ct),
                (parentId, nextLink) => GetUserChildPageAsync(parentId, nextLink, ct));

    public async Task<FolderInfo> CreateFolderAsync(
        string displayName, string? parentFolderId, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(displayName);
        string name = displayName.Trim();
        string? parentId = string.IsNullOrWhiteSpace(parentFolderId) ? null : parentFolderId.Trim();

        var folder = new MailFolder { DisplayName = name };

        MailFolder? created;
        if (IsMe)
        {
            created = parentId is null
                ? await client.Me.MailFolders.PostAsync(folder, cancellationToken: ct).ConfigureAwait(false)
                : await client.Me.MailFolders[parentId].ChildFolders
                    .PostAsync(folder, cancellationToken: ct).ConfigureAwait(false);
        }
        else
        {
            created = parentId is null
                ? await client.Users[_options.UserIdOrUpn].MailFolders
                    .PostAsync(folder, cancellationToken: ct).ConfigureAwait(false)
                : await client.Users[_options.UserIdOrUpn].MailFolders[parentId].ChildFolders
                    .PostAsync(folder, cancellationToken: ct).ConfigureAwait(false);
        }

        return created is null
            ? throw GraphServiceException.GraphError(0, null, "Folder creation returned no result.")
            : MapFolder(created, parentId, parentId is null ? name : null);
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
                ? throw GraphServiceException.MessageNotFound(messageId, "outlook_move_email")
                : EmailMapper.MapSummary(moved);
        }

        var userBody = new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody { DestinationId = destId };
        var userMoved = await client.Users[_options.UserIdOrUpn].Messages[messageId].Move
            .PostAsync(userBody, cancellationToken: ct).ConfigureAwait(false);
        return userMoved is null
            ? throw GraphServiceException.MessageNotFound(messageId, "outlook_move_email")
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
        RecipientGuard.ValidateRecipients(to, _policy);
        body = MessageDisclosure.Apply(body, isHtml, _policy);

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
            ? throw GraphServiceException.GraphError(0, null, "Draft creation returned no result.")
            : EmailMapper.MapDetail(draft);
    }

    public async Task<EmailDetail> CreateReplyDraftAsync(
        string messageId, string comment, CancellationToken ct = default)
    {
        RequireId(messageId);
        RequireBody(comment);

        if (_policy.RequireInternalRecipients)
        {
            // Replies go to Reply-To when set, else to From: resolve server-side
            // (not from LLM input) and enforce the domain policy on the targets.
            // Missing sender fails closed — never fail open.
            IReadOnlyList<string> targets = await GetReplyTargetsAsync(messageId, ct).ConfigureAwait(false);
            if (targets.Count == 0)
            {
                throw GraphServiceException.InvalidRequest(
                    $"Cannot verify the reply recipient of message '{messageId}' (no sender address).",
                    "pick a mail with a sender address, or ask your admin about policy.json");
            }

            RecipientGuard.ValidateRecipients(targets, _policy);
        }

        comment = MessageDisclosure.ApplyAuto(comment, _policy);

        if (IsMe)
        {
            var body = new Microsoft.Graph.Me.Messages.Item.CreateReply.CreateReplyPostRequestBody
            {
                Comment = comment
            };
            var draft = await client.Me.Messages[messageId].CreateReply
                .PostAsync(body, cancellationToken: ct).ConfigureAwait(false);
            return draft is null
                ? throw GraphServiceException.MessageNotFound(messageId, "outlook_create_reply_draft")
                : EmailMapper.MapDetail(draft);
        }

        var userBody = new Microsoft.Graph.Users.Item.Messages.Item.CreateReply.CreateReplyPostRequestBody
        {
            Comment = comment
        };
        var userDraft = await client.Users[_options.UserIdOrUpn].Messages[messageId].CreateReply
            .PostAsync(userBody, cancellationToken: ct).ConfigureAwait(false);
        return userDraft is null
            ? throw GraphServiceException.MessageNotFound(messageId, "outlook_create_reply_draft")
            : EmailMapper.MapDetail(userDraft);
    }

    public async Task<EmailDetail> CreateForwardDraftAsync(
        string messageId, IReadOnlyList<string> to, string? comment = null, CancellationToken ct = default)
    {
        RequireId(messageId);
        RequireRecipients(to);
        RecipientGuard.ValidateRecipients(to, _policy);

        comment = comment is null ? null : MessageDisclosure.ApplyAuto(comment, _policy);
        comment ??= _policy.AiDisclosureEnabled && !string.IsNullOrWhiteSpace(_policy.AiDisclosureText)
            ? MessageDisclosure.Apply(string.Empty, isHtml: false, _policy)
            : null;

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
                ? throw GraphServiceException.MessageNotFound(messageId, "outlook_create_forward_draft")
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
            ? throw GraphServiceException.MessageNotFound(messageId, "outlook_create_forward_draft")
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
            throw GraphServiceException.InvalidRequest(
                "Nothing to update: provide at least one of subject, body or to.",
                "pass subject and/or body and/or to, then retry");
        }

        var patch = new Message();
        if (subject is not null)
        {
            patch.Subject = subject;
        }

        // Policy-relevant current state, fetched once when needed (raw message:
        // EmailDetail.Body is truncated and must never be written back).
        Message? rawForPolicy = null;
        async Task<Message> RawForPolicyAsync() =>
            rawForPolicy ??= await GetRawMessageAsync(messageId, ct).ConfigureAwait(false);

        if (to is not null)
        {
            RecipientGuard.ValidateRecipients(to, _policy);
            patch.ToRecipients = [.. to.Select(ToRecipient)];
        }
        else if (_policy.RequireInternalRecipients)
        {
            // Without replacement recipients Graph keeps the draft's existing ones:
            // validate those server-side instead of skipping the check.
            Message raw = await RawForPolicyAsync().ConfigureAwait(false);
            RecipientGuard.ValidateRecipients(
                [.. (raw.ToRecipients ?? []).Select(r => r.EmailAddress?.Address)
                    .OfType<string>().Where(a => !string.IsNullOrWhiteSpace(a))],
                _policy);
        }

        if (body is not null)
        {
            patch.Body = new ItemBody
            {
                ContentType = isHtml ? BodyType.Html : BodyType.Text,
                Content = MessageDisclosure.Apply(body, isHtml, _policy)
            };
        }
        else if (_policy.AiDisclosureEnabled && !string.IsNullOrWhiteSpace(_policy.AiDisclosureText))
        {
            // Subject/recipient-only updates must not leave a disclosure-less body behind.
            Message raw = await RawForPolicyAsync().ConfigureAwait(false);
            string currentBody = raw.Body?.Content ?? string.Empty;
            bool currentIsHtml = raw.Body?.ContentType == BodyType.Html;
            string updated = MessageDisclosure.Apply(currentBody, currentIsHtml, _policy);
            if (updated != currentBody)
            {
                patch.Body = new ItemBody
                {
                    ContentType = raw.Body?.ContentType ?? BodyType.Text,
                    Content = updated
                };
            }
        }

        if (IsMe)
        {
            var updated = await client.Me.Messages[messageId]
                .PatchAsync(patch, cancellationToken: ct).ConfigureAwait(false);
            return updated is null
                ? throw GraphServiceException.MessageNotFound(messageId, "outlook_update_draft")
                : EmailMapper.MapDetail(updated);
        }

        var userUpdated = await client.Users[_options.UserIdOrUpn].Messages[messageId]
            .PatchAsync(patch, cancellationToken: ct).ConfigureAwait(false);
        return userUpdated is null
            ? throw GraphServiceException.MessageNotFound(messageId, "outlook_update_draft")
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
            throw GraphServiceException.MissingId("attachmentId");
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
            throw GraphServiceException.InvalidRequest(
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
            throw GraphServiceException.AttachmentTooLarge(file.Name ?? "?", bytes.Length, cap);
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
                ? throw GraphServiceException.MessageNotFound(messageId, "outlook_set_categories")
                : EmailMapper.MapDetail(updated);
        }

        var userUpdated = await client.Users[_options.UserIdOrUpn].Messages[messageId]
            .PatchAsync(patch, cancellationToken: ct).ConfigureAwait(false);
        return userUpdated is null
            ? throw GraphServiceException.MessageNotFound(messageId, "outlook_set_categories")
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
            throw GraphServiceException.InvalidRequest(
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

    private async Task<IReadOnlyList<string>> GetReplyTargetsAsync(string messageId, CancellationToken ct)
    {
        Message raw = await GetRawMessageAsync(messageId, ct).ConfigureAwait(false);

        var targets = new List<string>();
        if (raw.ReplyTo is { Count: > 0 })
        {
            targets.AddRange(raw.ReplyTo
                .Select(r => r.EmailAddress?.Address)
                .OfType<string>()
                .Where(a => !string.IsNullOrWhiteSpace(a)));
        }
        else if (!string.IsNullOrWhiteSpace(raw.From?.EmailAddress?.Address))
        {
            targets.Add(raw.From!.EmailAddress!.Address!);
        }

        return targets;
    }

    /// <summary>Raw message with the policy-relevant fields (untruncated body).
    /// Used only for server-side policy checks, never returned to the LLM.</summary>
    private async Task<Message> GetRawMessageAsync(string messageId, CancellationToken ct)
    {
        RequireId(messageId);
        Message? msg = IsMe
            ? await client.Me.Messages[messageId].GetAsync(c =>
                c.QueryParameters.Select = ["from", "replyTo", "toRecipients", "body"], ct).ConfigureAwait(false)
            : await client.Users[_options.UserIdOrUpn].Messages[messageId].GetAsync(c =>
                c.QueryParameters.Select = ["from", "replyTo", "toRecipients", "body"], ct).ConfigureAwait(false);

        return msg is null
            ? throw GraphServiceException.MessageNotFound(messageId, "policy-check")
            : msg;
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

    private async Task<IReadOnlyList<FolderInfo>> ListFoldersAsync(
        Func<string?, Task<FolderPage>> getRootPage,
        Func<string, string?, Task<FolderPage>> getChildPage)
    {
        var result = new List<FolderInfo>();
        var queue = new Queue<FolderNode>();
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        string? nextLink = null;
        do
        {
            FolderPage page = await getRootPage(nextLink).ConfigureAwait(false);
            foreach (MailFolder folder in page.Items)
            {
                if (!TryVisit(folder, visited, out string folderId))
                {
                    continue;
                }

                string path = folder.DisplayName ?? string.Empty;
                result.Add(MapFolder(folder, folder.ParentFolderId, path));
                EnqueueIfHasChildren(queue, folder, folderId, path);
            }

            nextLink = page.NextLink;
        }
        while (!string.IsNullOrWhiteSpace(nextLink));

        while (queue.Count > 0)
        {
            FolderNode parent = queue.Dequeue();
            nextLink = null;
            do
            {
                FolderPage page = await getChildPage(parent.Id, nextLink).ConfigureAwait(false);
                foreach (MailFolder folder in page.Items)
                {
                    if (!TryVisit(folder, visited, out string folderId))
                    {
                        continue;
                    }

                    string path = AppendPath(parent.Path, folder.DisplayName);
                    string parentId = folder.ParentFolderId ?? parent.Id;
                    result.Add(MapFolder(folder, parentId, path));
                    EnqueueIfHasChildren(queue, folder, folderId, path);
                }

                nextLink = page.NextLink;
            }
            while (!string.IsNullOrWhiteSpace(nextLink));
        }

        return result;
    }

    private async Task<FolderPage> GetMeRootPageAsync(string? nextLink, CancellationToken ct)
    {
        var page = string.IsNullOrWhiteSpace(nextLink)
            ? await client.Me.MailFolders.GetAsync(c =>
            {
                c.QueryParameters.Top = 100;
                c.QueryParameters.Select = FolderSelect;
            }, ct).ConfigureAwait(false)
            : await client.Me.MailFolders.WithUrl(nextLink!).GetAsync(cancellationToken: ct).ConfigureAwait(false);
        return ToFolderPage(page);
    }

    private async Task<FolderPage> GetMeChildPageAsync(
        string parentId, string? nextLink, CancellationToken ct)
    {
        var page = string.IsNullOrWhiteSpace(nextLink)
            ? await client.Me.MailFolders[parentId].ChildFolders.GetAsync(c =>
            {
                c.QueryParameters.Top = 100;
                c.QueryParameters.Select = FolderSelect;
            }, ct).ConfigureAwait(false)
            : await client.Me.MailFolders[parentId].ChildFolders
                .WithUrl(nextLink!).GetAsync(cancellationToken: ct).ConfigureAwait(false);
        return ToFolderPage(page);
    }

    private async Task<FolderPage> GetUserRootPageAsync(string? nextLink, CancellationToken ct)
    {
        var folders = client.Users[_options.UserIdOrUpn].MailFolders;
        var page = string.IsNullOrWhiteSpace(nextLink)
            ? await folders.GetAsync(c =>
            {
                c.QueryParameters.Top = 100;
                c.QueryParameters.Select = FolderSelect;
            }, ct).ConfigureAwait(false)
            : await folders.WithUrl(nextLink!).GetAsync(cancellationToken: ct).ConfigureAwait(false);
        return ToFolderPage(page);
    }

    private async Task<FolderPage> GetUserChildPageAsync(
        string parentId, string? nextLink, CancellationToken ct)
    {
        var childFolders = client.Users[_options.UserIdOrUpn].MailFolders[parentId].ChildFolders;
        var page = string.IsNullOrWhiteSpace(nextLink)
            ? await childFolders.GetAsync(c =>
            {
                c.QueryParameters.Top = 100;
                c.QueryParameters.Select = FolderSelect;
            }, ct).ConfigureAwait(false)
            : await childFolders.WithUrl(nextLink!).GetAsync(cancellationToken: ct).ConfigureAwait(false);
        return ToFolderPage(page);
    }

    private static FolderPage ToFolderPage(Microsoft.Graph.Models.MailFolderCollectionResponse? page) =>
        new([.. page?.Value ?? []], page?.OdataNextLink);

    private static bool TryVisit(MailFolder folder, ISet<string> visited, out string folderId)
    {
        folderId = folder.Id ?? string.Empty;
        return !string.IsNullOrWhiteSpace(folderId) && visited.Add(folderId);
    }

    private static void EnqueueIfHasChildren(
        Queue<FolderNode> queue,
        MailFolder folder,
        string folderId,
        string path)
    {
        if (folder.ChildFolderCount is not 0)
        {
            queue.Enqueue(new FolderNode(folderId, path));
        }
    }

    private static string AppendPath(string parentPath, string? displayName) =>
        string.IsNullOrWhiteSpace(parentPath)
            ? displayName ?? string.Empty
            : $"{parentPath}/{displayName ?? string.Empty}";

    private static void RequireId(string messageId)
    {
        if (string.IsNullOrWhiteSpace(messageId))
        {
            throw GraphServiceException.MissingId();
        }
    }

    private static void RequireRecipients(IReadOnlyList<string> to)
    {
        if (to.Count == 0 || to.Any(string.IsNullOrWhiteSpace))
        {
            throw GraphServiceException.InvalidRequest(
                "At least one valid recipient address is required.",
                "pass non-empty 'to' addresses, e.g. [\"a@example.com\"]");
        }
    }

    private static void RequireBody(string? body)
    {
        if (string.IsNullOrWhiteSpace(body))
        {
            throw GraphServiceException.InvalidRequest(
                "Body must not be empty.",
                "provide the mail text (or html with isHtml=true)");
        }
    }

    private static FolderInfo MapFolder(MailFolder f, string? parentId, string? path) => new(
        f.Id ?? string.Empty,
        f.DisplayName ?? string.Empty,
        f.TotalItemCount ?? 0,
        f.UnreadItemCount ?? 0,
        f.ParentFolderId ?? parentId,
        path);

    private sealed record FolderPage(IReadOnlyList<MailFolder> Items, string? NextLink);

    private sealed record FolderNode(string Id, string Path);

    private static Recipient ToRecipient(string address) => new()
    {
        EmailAddress = new EmailAddress { Address = address }
    };
}
