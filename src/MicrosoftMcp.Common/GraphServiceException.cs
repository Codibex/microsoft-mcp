namespace MicrosoftMcp.Common;

/// <summary>
/// Agent-facing error with a machine-readable <see cref="Code"/> and a "Next:" hint
/// telling the caller how to recover. Message format: <c>[code] summary Next: hint</c>.
/// </summary>
public sealed class GraphServiceException : Exception
{
    public string Code { get; }

    private GraphServiceException(string code, string message, Exception? inner = null)
        : base(message, inner)
    {
        Code = code;
    }

    private static GraphServiceException Create(string code, string summary, string hint, Exception? inner = null) =>
        new(code, $"[{code}] {summary} Next: {hint}", inner);

    public static GraphServiceException MessageNotFound(string messageId, string operation) => Create(
        "message-not-found",
        $"Message '{messageId}' was not found during {operation}.",
        "call outlook_search_emails to get a valid Graph id (use 'id', not internetMessageId; ids expire after moves)");

    public static GraphServiceException FolderNotFound(string destination) => Create(
        "folder-not-found",
        $"Folder '{destination}' was not found.",
        "call outlook_list_folders and use id, displayName or a well-known name (inbox, archive, deleteditems, drafts)");

    public static GraphServiceException DriveItemNotFound(string itemRef, string operation) => Create(
        "item-not-found",
        $"Drive item '{itemRef}' was not found during {operation}.",
        "call onedrive_list_children or onedrive_search_files to get a valid id or /path (ids change on move)");

    public static GraphServiceException CalendarNotFound(string calendarId, string operation) => Create(
        "calendar-not-found",
        $"Calendar '{calendarId}' was not found during {operation}.",
        "call calendar_list_calendars to get valid calendar ids");

    public static GraphServiceException EventNotFound(string eventId, string operation) => Create(
        "event-not-found",
        $"Event '{eventId}' was not found during {operation}.",
        "call calendar_list_events for the time window to get valid event ids (ids change on move)");

    public static GraphServiceException TeamNotFound(string teamId, string operation) => Create(
        "team-not-found",
        $"Team '{teamId}' was not found during {operation}.",
        "call teams_list_teams to get valid team ids");

    public static GraphServiceException ChannelNotFound(string channelId, string operation) => Create(
        "channel-not-found",
        $"Channel '{channelId}' was not found during {operation}.",
        "call teams_list_channels for the team to get valid channel ids");

    public static GraphServiceException ChatNotFound(string chatId, string operation) => Create(
        "chat-not-found",
        $"Chat '{chatId}' was not found during {operation}.",
        "call teams_list_chats to get valid chat ids");

    public static GraphServiceException TeamsMessageNotFound(string messageId, string operation) => Create(
        "channel-message-not-found",
        $"Message '{messageId}' was not found during {operation}.",
        "call teams_list_channel_messages or teams_list_chat_messages to get valid message ids");

    public static GraphServiceException MeetingTranscriptNotFound(string transcriptId, string operation) => Create(
        "meeting-transcript-not-found",
        $"Meeting transcript '{transcriptId}' was not found during {operation}.",
        "call teams_list_meeting_transcripts with the online meeting id to get valid transcript ids");

    public static GraphServiceException MeetingInsightNotFound(string insightId, string operation) => Create(
        "meeting-insight-not-found",
        $"Meeting insight '{insightId}' was not found during {operation}.",
        "call teams_list_meeting_insights with the online meeting id to get valid insight ids");

    public static GraphServiceException MissingRef(string what = "itemRef") => Create(
        "invalid-request",
        $"{what} must not be empty.",
        "use \"root\", an item id, or a /path/from/root");

    public static GraphServiceException InvalidRequest(string detail, string hint) => Create(
        "invalid-request", detail, hint);

    public static GraphServiceException MissingId(string what = "messageId") => Create(
        "invalid-request",
        $"{what} must not be empty.",
        "use an id from outlook_search_emails or outlook_read_email");

    public static GraphServiceException AuthMisconfigured(string detail) => Create(
        "auth-misconfigured",
        detail,
        "set Graph__TenantId and Graph__ClientId (delegated) or additionally Graph__UserIdOrUpn and Graph__ClientSecret (app-only)");

    public static GraphServiceException AuthFailed(string detail, string? hint = null, Exception? inner = null) => Create(
        "auth-failed",
        detail,
        hint ?? "re-authenticate (delete the token cache), verify TenantId/ClientId, and for headless hosts use DelegatedFlow=DeviceCode",
        inner);

    public static GraphServiceException AuthCacheUnavailable(Exception? inner = null) => Create(
        "auth-cache-unavailable",
        "Persistent token cache unavailable; the Linux Secret Service/GNOME Keyring is not reachable.",
        "install and start a Secret Service, set Graph__FallbackToMemoryTokenCache=true, or disable persistence with Graph__EnableTokenCache=false",
        inner);

    public static GraphServiceException AccessDenied(int status, string? graphCode, string? detail) => Create(
        "access-denied",
        $"Graph denied access (HTTP {status}{(graphCode is null ? string.Empty : $", {graphCode}")}). {detail ?? "No detail."}",
        "verify Entra consent and the Graph permission required by this operation; application permissions require admin consent");

    public static GraphServiceException MailboxUnavailable(string? graphCode, string? detail) => Create(
        "mailbox-unavailable",
        $"The mailbox is not usable via Graph{(graphCode is null ? string.Empty : $" ({graphCode})")}. {detail ?? "No detail."}",
        "sign in with the account that owns an active Exchange Online mailbox (check license, no soft-delete/on-prem); delegated users must not use an admin-only account without mailbox");

    public static GraphServiceException Throttled(string? detail) => Create(
        "throttled",
        $"Graph rate-limited the request (429). {detail ?? string.Empty}".Trim(),
        "wait ~60s, then retry with a smaller 'top' value");

    public static GraphServiceException AttachmentTooLarge(string name, int size, int maxBytes) => Create(
        "attachment-too-large",
        $"Attachment '{name}' is {size} bytes, above the {maxBytes} byte limit.",
        "pick a smaller attachment (see size in outlook_list_attachments), or raise maxBytes up to 2097152");

    public static GraphServiceException Conflict(string operation, string? detail) => Create(
        "conflict",
        $"Concurrent change during {operation} (409). {detail ?? string.Empty}".Trim(),
        "re-read the message (outlook_read_email) to get fresh state, then retry");

    public static GraphServiceException ServiceUnavailable(string? detail, Exception? inner = null) => Create(
        "service-unavailable",
        $"Graph is unreachable or errored. {detail ?? string.Empty}".Trim(),
        "retry once after a short wait; if persistent, check network/proxy and Entra app status",
        inner);

    public static GraphServiceException GraphError(int status, string? graphCode, string? detail) => Create(
        "graph-error",
        $"Unexpected Graph error (HTTP {status}{(graphCode is null ? string.Empty : $", {graphCode}")}). {detail ?? "No detail."}",
        "retry once; if persistent, narrow the request (smaller top, valid folder) and report status + graph code");
}
