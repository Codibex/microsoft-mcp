namespace MicrosoftMcp.Outlook.Tests;

using MicrosoftMcp.Common;

/// <summary>In-memory fake of <see cref="IGraphMailService"/> for tool-level tests. No Graph, no network.</summary>
internal sealed class FakeGraphMailService : IGraphMailService
{
    private sealed class Stored
    {
        public required string Id { get; init; }
        public string Subject { get; set; } = string.Empty;
        public string From { get; init; } = "boss@example.com";
        public string Folder { get; set; } = "inbox";
        public bool IsRead { get; set; }
        public bool HasAttachments { get; init; }
        public List<string> Categories { get; set; } = [];
        public string Importance { get; set; } = "normal";
        public string Body { get; set; } = "body";
    }

    private readonly Dictionary<string, Stored> _messages = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, FolderInfo> _folders = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<AttachmentInfo> _attachments =
    [
        new("a1", "rechnung.pdf", "application/pdf", 1234, false)
    ];

    private readonly List<CategoryInfo> _categories =
    [
        new("c-red", "Red category", "preset0"),
        new("c-blue", "Blue category", "preset1")
    ];

    private int _drafts;

    public EmailQuery? LastQuery { get; private set; }

    public FakeGraphMailService()
    {
        foreach (var (id, name) in new[] { ("inbox", "Inbox"), ("archive", "Archive"),
                     ("deleteditems", "DeletedItems"), ("drafts", "Drafts"), ("f-projekte", "Projekte") })
        {
            _folders[id] = new FolderInfo(id, name, 0, 0);
        }

        Seed(new Stored { Id = "m1", Subject = "Rechnung Januar", IsRead = false, HasAttachments = true });
        Seed(new Stored { Id = "m2", Subject = "Hallo", IsRead = true, Categories = ["Red category"] });
        Seed(new Stored { Id = "m3", Subject = "Alt", Folder = "archive", IsRead = true });
    }

    private void Seed(Stored s)
    {
        _messages[s.Id] = s;
        RefreshCounts();
    }

    private Stored Get(string id)
    {
        if (string.IsNullOrWhiteSpace(id))
        {
            throw MailServiceException.MissingId();
        }

        return _messages.TryGetValue(id, out var m)
            ? m
            : throw MailServiceException.MessageNotFound(id, "fake");
    }

    private static EmailSummary ToSummary(Stored m) => new(
        m.Id, m.Subject, new EmailAddressDto("Boss", m.From), [],
        DateTimeOffset.UtcNow, m.IsRead, m.HasAttachments, m.Categories,
        m.Importance, "preview");

    private static EmailDetail ToDetail(Stored m) => new(
        m.Id, m.Subject, new EmailAddressDto("Boss", m.From), [],
        DateTimeOffset.UtcNow, m.IsRead, m.Categories, "preview", m.Body, null);

    private void RefreshCounts()
    {
        foreach (var key in _folders.Keys)
        {
            var inFolder = _messages.Values.Where(m => m.Folder == key).ToList();
            _folders[key] = _folders[key] with
            {
                TotalCount = inFolder.Count,
                UnreadCount = inFolder.Count(m => !m.IsRead)
            };
        }
    }

    public Task<IReadOnlyList<EmailSummary>> SearchAsync(EmailQuery query, CancellationToken ct = default)
    {
        LastQuery = query;
        IEnumerable<Stored> result = _messages.Values;

        if (!string.IsNullOrWhiteSpace(query.Folder))
        {
            string folderId = FolderResolver.Resolve(query.Folder!, [.. _folders.Values]);
            result = result.Where(m => m.Folder == folderId);
        }

        if (!string.IsNullOrWhiteSpace(query.Query))
        {
            result = result.Where(m => m.Subject.Contains(query.Query!, StringComparison.OrdinalIgnoreCase));
        }

        if (!string.IsNullOrWhiteSpace(query.From))
        {
            result = result.Where(m => m.From.Contains(query.From!, StringComparison.OrdinalIgnoreCase));
        }

        int top = Math.Clamp(query.Top, 1, 50);
        return Task.FromResult<IReadOnlyList<EmailSummary>>([.. result.Take(top).Select(ToSummary)]);
    }

    public Task<EmailDetail> GetAsync(string messageId, CancellationToken ct = default) =>
        Task.FromResult(ToDetail(Get(messageId)));

    public Task<IReadOnlyList<FolderInfo>> ListFoldersAsync(CancellationToken ct = default) =>
        Task.FromResult<IReadOnlyList<FolderInfo>>([.. _folders.Values]);

    public Task<FolderInfo> CreateFolderAsync(string displayName, string? parentFolderId = null, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(displayName);
        string id = "f-" + displayName.Trim().ToLowerInvariant().Replace(' ', '-');
        var folder = new FolderInfo(id, displayName.Trim(), 0, 0);
        _folders[id] = folder;
        return Task.FromResult(folder);
    }

    public Task<EmailSummary> MoveAsync(string messageId, string destination, CancellationToken ct = default)
    {
        var msg = Get(messageId);
        msg.Folder = FolderResolver.Resolve(destination, [.. _folders.Values]);
        RefreshCounts();
        return Task.FromResult(ToSummary(msg));
    }

    public Task<EmailSummary> ArchiveAsync(string messageId, CancellationToken ct = default) =>
        MoveAsync(messageId, "archive", ct);

    public Task DeleteAsync(string messageId, CancellationToken ct = default) =>
        MoveAsync(messageId, "deleteditems", ct).ContinueWith(_ => { }, ct);

    public Task<EmailDetail> CreateDraftAsync(
        IReadOnlyList<string> to, string subject, string body, bool isHtml = false, CancellationToken ct = default)
    {
        RequireRecipients(to);
        var draft = new Stored
        {
            Id = $"draft-{++_drafts}",
            Subject = subject,
            Body = body,
            Folder = "drafts"
        };
        Seed(draft);
        return Task.FromResult(ToDetail(draft));
    }

    public Task<EmailDetail> CreateReplyDraftAsync(string messageId, string comment, CancellationToken ct = default)
    {
        var orig = Get(messageId);
        var draft = new Stored
        {
            Id = $"draft-{++_drafts}",
            Subject = "Re: " + orig.Subject,
            Body = comment,
            Folder = "drafts"
        };
        Seed(draft);
        return Task.FromResult(ToDetail(draft));
    }

    public Task<EmailDetail> CreateForwardDraftAsync(
        string messageId, IReadOnlyList<string> to, string? comment = null, CancellationToken ct = default)
    {
        RequireRecipients(to);
        var orig = Get(messageId);
        var draft = new Stored
        {
            Id = $"draft-{++_drafts}",
            Subject = "Fwd: " + orig.Subject,
            Body = (comment ?? string.Empty) + "\n---\n" + orig.Body,
            Folder = "drafts"
        };
        Seed(draft);
        return Task.FromResult(ToDetail(draft));
    }

    public Task<EmailDetail> UpdateDraftAsync(
        string messageId, string? subject = null, string? body = null,
        bool isHtml = false, IReadOnlyList<string>? to = null, CancellationToken ct = default)
    {
        var draft = Get(messageId);
        if (subject is not null)
        {
            draft.Subject = subject;
        }

        if (body is not null)
        {
            draft.Body = body;
        }

        return Task.FromResult(ToDetail(draft));
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

    public Task<IReadOnlyList<AttachmentInfo>> ListAttachmentsAsync(string messageId, CancellationToken ct = default)
    {
        var msg = Get(messageId);
        IReadOnlyList<AttachmentInfo> result = msg.HasAttachments ? _attachments : [];
        return Task.FromResult(result);
    }

    public Task<IReadOnlyList<CategoryInfo>> ListCategoriesAsync(CancellationToken ct = default) =>
        Task.FromResult<IReadOnlyList<CategoryInfo>>(_categories);

    public Task<EmailDetail> SetCategoriesAsync(string messageId, string[] add, string[] remove, CancellationToken ct = default)
    {
        var msg = Get(messageId);
        msg.Categories = EmailMapper.MergeCategories(msg.Categories, add, remove);
        return Task.FromResult(ToDetail(msg));
    }

    public Task MarkReadAsync(string messageId, bool isRead, CancellationToken ct = default)
    {
        Get(messageId).IsRead = isRead;
        RefreshCounts();
        return Task.CompletedTask;
    }

    public Task SetImportanceAsync(string messageId, string importance, CancellationToken ct = default)
    {
        string lower = importance.ToLowerInvariant();
        if (lower is not ("low" or "normal" or "high"))
        {
            throw MailServiceException.InvalidRequest(
                $"Invalid importance '{importance}'.",
                "use low, normal or high");
        }

        Get(messageId).Importance = lower;
        return Task.CompletedTask;
    }
}
