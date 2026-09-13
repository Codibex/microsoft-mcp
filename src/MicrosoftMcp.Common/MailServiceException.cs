namespace MicrosoftMcp.Common;

/// <summary>
/// Agent-facing error with a machine-readable <see cref="Code"/> and a "Next:" hint
/// telling the caller how to recover. Message format: <c>[code] summary Next: hint</c>.
/// </summary>
public sealed class MailServiceException : Exception
{
    public string Code { get; }

    private MailServiceException(string code, string message, Exception? inner = null)
        : base(message, inner)
    {
        Code = code;
    }

    private static MailServiceException Create(string code, string summary, string hint, Exception? inner = null) =>
        new(code, $"[{code}] {summary} Next: {hint}", inner);

    public static MailServiceException MessageNotFound(string messageId, string operation) => Create(
        "message-not-found",
        $"Message '{messageId}' was not found during {operation}.",
        "call search_emails to get a valid Graph id (use 'id', not internetMessageId; ids expire after moves)");

    public static MailServiceException FolderNotFound(string destination) => Create(
        "folder-not-found",
        $"Folder '{destination}' was not found.",
        "call list_folders and use id, displayName or a well-known name (inbox, archive, deleteditems, drafts)");

    public static MailServiceException DriveItemNotFound(string itemRef, string operation) => Create(
        "item-not-found",
        $"Drive item '{itemRef}' was not found during {operation}.",
        "call list_children or search_files to get a valid id or /path (ids change on move)");

    public static MailServiceException CalendarNotFound(string calendarId, string operation) => Create(
        "calendar-not-found",
        $"Calendar '{calendarId}' was not found during {operation}.",
        "call list_calendars to get valid calendar ids");

    public static MailServiceException EventNotFound(string eventId, string operation) => Create(
        "event-not-found",
        $"Event '{eventId}' was not found during {operation}.",
        "call list_events for the time window to get valid event ids (ids change on move)");

    public static MailServiceException TeamNotFound(string teamId, string operation) => Create(
        "team-not-found",
        $"Team '{teamId}' was not found during {operation}.",
        "call list_teams to get valid team ids");

    public static MailServiceException ChannelNotFound(string channelId, string operation) => Create(
        "channel-not-found",
        $"Channel '{channelId}' was not found during {operation}.",
        "call list_channels for the team to get valid channel ids");

    public static MailServiceException ChatNotFound(string chatId, string operation) => Create(
        "chat-not-found",
        $"Chat '{chatId}' was not found during {operation}.",
        "call list_chats to get valid chat ids");

    public static MailServiceException TeamsMessageNotFound(string messageId, string operation) => Create(
        "channel-message-not-found",
        $"Message '{messageId}' was not found during {operation}.",
        "call list_channel_messages or list_chat_messages to get valid message ids");

    public static MailServiceException MissingRef(string what = "itemRef") => Create(
        "invalid-request",
        $"{what} must not be empty.",
        "use \"root\", an item id, or a /path/from/root");

    public static MailServiceException InvalidRequest(string detail, string hint) => Create(
        "invalid-request", detail, hint);

    public static MailServiceException MissingId(string what = "messageId") => Create(
        "invalid-request",
        $"{what} must not be empty.",
        "use an id from search_emails or read_email");

    public static MailServiceException AuthMisconfigured(string detail) => Create(
        "auth-misconfigured",
        detail,
        "set Graph__TenantId and Graph__ClientId (delegated) or additionally Graph__UserIdOrUpn and Graph__ClientSecret (app-only)");

    public static MailServiceException AuthFailed(string detail, string? hint = null, Exception? inner = null) => Create(
        "auth-failed",
        detail,
        hint ?? "re-authenticate (delete the token cache), verify TenantId/ClientId, and for headless hosts use DelegatedFlow=DeviceCode",
        inner);

    public static MailServiceException AccessDenied(int status, string? graphCode, string? detail) => Create(
        "access-denied",
        $"Graph denied access (HTTP {status}{(graphCode is null ? string.Empty : $", {graphCode}")}). {detail ?? "No detail."}",
        "verify Entra consent and scopes (delegated: Mail.Read/Mail.ReadWrite; app-only: Mail.ReadWrite + admin consent)");

    public static MailServiceException MailboxUnavailable(string? graphCode, string? detail) => Create(
        "mailbox-unavailable",
        $"The mailbox is not usable via Graph{(graphCode is null ? string.Empty : $" ({graphCode})")}. {detail ?? "No detail."}",
        "sign in with the account that owns an active Exchange Online mailbox (check license, no soft-delete/on-prem); delegated users must not use an admin-only account without mailbox");

    public static MailServiceException Throttled(string? detail) => Create(
        "throttled",
        $"Graph rate-limited the request (429). {detail ?? string.Empty}".Trim(),
        "wait ~60s, then retry with a smaller 'top' value");

    public static MailServiceException AttachmentTooLarge(string name, int size, int maxBytes) => Create(
        "attachment-too-large",
        $"Attachment '{name}' is {size} bytes, above the {maxBytes} byte limit.",
        "pick a smaller attachment (see size in list_attachments), or raise maxBytes up to 2097152");

    public static MailServiceException Conflict(string operation, string? detail) => Create(
        "conflict",
        $"Concurrent change during {operation} (409). {detail ?? string.Empty}".Trim(),
        "re-read the message (read_email) to get fresh state, then retry");

    public static MailServiceException ServiceUnavailable(string? detail, Exception? inner = null) => Create(
        "service-unavailable",
        $"Graph is unreachable or errored. {detail ?? string.Empty}".Trim(),
        "retry once after a short wait; if persistent, check network/proxy and Entra app status",
        inner);

    public static MailServiceException GraphError(int status, string? graphCode, string? detail) => Create(
        "graph-error",
        $"Unexpected Graph error (HTTP {status}{(graphCode is null ? string.Empty : $", {graphCode}")}). {detail ?? "No detail."}",
        "retry once; if persistent, narrow the request (smaller top, valid folder) and report status + graph code");
}
