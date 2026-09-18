using System.Text.RegularExpressions;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Kiota.Abstractions;

namespace MicrosoftMcp.Common;

/// <summary>Translates Graph/Kiota/transport exceptions into agent-actionable errors.</summary>
public static partial class GraphErrorMapper
{
    private static readonly Dictionary<string, Func<int, string?, string, string, GraphServiceException>> GraphCodeOverrides =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["ErrorItemNotFound"] = static (_, _, op, res) => NotFoundFor(res, op),
            ["ErrorMessageNotFound"] = static (_, _, op, res) => NotFoundFor(res, op),
            ["ErrorFolderNotFound"] = static (_, _, _, _) =>
                GraphServiceException.FolderNotFound("<unknown>"),
            ["ErrorAccessDenied"] = static (s, c, _, _) =>
                GraphServiceException.AccessDenied(s, c, null),
            ["ErrorInvalidIdMalformed"] = static (_, _, _, _) =>
                GraphServiceException.InvalidRequest(
                    "Malformed Graph id.",
                    "use the 'id' field from the matching list/search tool, not a web link or secondary id"),
            ["MailboxNotEnabledForRESTAPI"] = static (_, c, d, _) =>
                GraphServiceException.MailboxUnavailable(c, d),
            ["ErrorMailboxNotAssociated"] = static (_, c, d, _) =>
                GraphServiceException.MailboxUnavailable(c, d),
            ["ErrorMailboxStoreUnknown"] = static (_, c, d, _) =>
                GraphServiceException.MailboxUnavailable(c, d),
            ["ErrorNonExistentMailbox"] = static (_, c, d, _) =>
                GraphServiceException.MailboxUnavailable(c, d),
            ["ErrorQuotaExceeded"] = static (_, _, _, _) =>
                GraphServiceException.InvalidRequest(
                    "Mailbox quota exceeded.",
                    "free mailbox space or narrow the operation before retrying"),
        };

    /// <summary>Known AADSTS codes with specific recovery hints (MSAL buries
    /// them behind generic sentences, so we promote code + hint).</summary>
    private static readonly Dictionary<string, string> AadStsHints =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["AADSTS700016"] = "the app is not visible in this directory yet: verify signInAudience covers this account type, wait a few minutes for propagation, then re-authenticate with the right account",
            ["AADSTS7000218"] = "enable 'Allow public client flows' in the app registration (Authentication page)",
            ["AADSTS65001"] = "consent is required: sign in with an account that can consent, or ask an admin to grant the Graph mail permissions",
            ["AADSTS50020"] = "this account does not exist in the target tenant: use the matching account, or switch TenantId to common/consumers",
            ["AADSTS90002"] = "TenantId is wrong: verify the directory ID, or use common/consumers",
            ["AADSTS2346"] = "this app is personal-only: use TenantId=consumers; use common only when signInAudience is AzureADandPersonalMicrosoftAccount",
        };

    /// <summary>Maps Graph/Kiota failures. Resource selects the not-found
    /// wording: "message" (mail), "calendar", "event", "drive" or "folder".</summary>
    public static GraphServiceException ToGraphServiceException(
        Exception ex, string operation, string resource = "message") =>
        ex switch
        {
            GraphServiceException already => already,
            ApiException api => FromStatus(
                api.ResponseStatusCode, ResolveGraphCode(api), Truncate(api.Message, 300), operation, resource),
            HttpRequestException http => GraphServiceException.ServiceUnavailable(
                Truncate(http.Message, 200), http),
            TimeoutException timeout => GraphServiceException.ServiceUnavailable(
                Truncate(timeout.Message, 200), timeout),
            ArgumentException arg => GraphServiceException.InvalidRequest(
                Truncate(arg.Message, 300) ?? "Invalid argument.",
                "check ids, folder names and enum values (low|normal|high), then retry"),
            { } other when IsAuthNamespace(other) => MapAuthFailure(AuthDetail(other), other),
            _ => GraphServiceException.GraphError(0, ex.GetType().Name, Truncate(ex.Message, 300))
        };

    private static GraphServiceException MapAuthFailure(string detail, Exception? ex)
    {
        string? hint = ExtractAadStsCode(detail) is { } code
            && AadStsHints.TryGetValue(code, out var specific)
            ? specific
            : null;
        return GraphServiceException.AuthFailed(detail, hint, ex);
    }

    public static GraphServiceException FromStatus(
        int status, string? graphCode, string? detail, string operation, string resource = "message")
    {
        if (graphCode is not null && GraphCodeOverrides.TryGetValue(graphCode, out var map))
        {
            return map(status, graphCode, operation, resource);
        }

        return status switch
        {
            400 => GraphServiceException.InvalidRequest(
                $"Graph rejected the request (400{(graphCode is null ? string.Empty : $", {graphCode}")}). {detail ?? string.Empty}".Trim(),
                "check ids, folder names and parameters, then retry with a narrower request"),
            401 => MapAuthFailure(
                $"Graph authentication failed (401{(graphCode is null ? string.Empty : $", {graphCode}")}). {detail ?? string.Empty}".Trim(),
                null),
            403 => GraphServiceException.AccessDenied(status, graphCode, detail),
            404 => NotFoundFor(resource, operation),
            409 => GraphServiceException.Conflict(operation, detail),
            429 => GraphServiceException.Throttled(detail),
            >= 500 => GraphServiceException.ServiceUnavailable(
                $"HTTP {status}{(graphCode is null ? string.Empty : $", {graphCode}")}. {detail ?? string.Empty}".Trim()),
            _ => GraphServiceException.GraphError(status, graphCode, detail)
        };
    }

    internal static GraphServiceException NotFoundFor(string resource, string operation) =>
        resource switch
        {
            "calendar" => GraphServiceException.CalendarNotFound("<unknown>", operation),
            "event" => GraphServiceException.EventNotFound("<unknown>", operation),
            "drive" => GraphServiceException.DriveItemNotFound("<unknown>", operation),
            "folder" => GraphServiceException.FolderNotFound("<unknown>"),
            "team" => GraphServiceException.TeamNotFound("<unknown>", operation),
            "channel" => GraphServiceException.ChannelNotFound("<unknown>", operation),
            "chat" => GraphServiceException.ChatNotFound("<unknown>", operation),
            "teams-message" => GraphServiceException.TeamsMessageNotFound("<unknown>", operation),
            "meeting-transcript" => GraphServiceException.MeetingTranscriptNotFound("<unknown>", operation),
            "meeting-insight" => GraphServiceException.MeetingInsightNotFound("<unknown>", operation),
            _ => GraphServiceException.MessageNotFound("<unknown>", operation)
        };

    /// <summary>Prefers the structured ODataError code; falls back to regex
    /// over the message text (Kiota puts the Graph body there).</summary>
    public static string? ResolveGraphCode(ApiException api) =>
        api is ODataError odata && !string.IsNullOrWhiteSpace(odata.Error?.Code)
            ? odata.Error.Code
            : ExtractGraphCode(api.Message);

    public static string? ExtractGraphCode(string? message)
    {
        if (string.IsNullOrEmpty(message))
        {
            return null;
        }

        var match = GraphCodeRegex().Match(message);
        return match.Success ? match.Groups[1].Value : null;
    }

    private static bool IsAuthNamespace(Exception ex) =>
        ex.GetType().FullName?.StartsWith("Azure.Identity.", StringComparison.Ordinal) == true
        || ex.GetType().FullName?.StartsWith("Microsoft.Identity.", StringComparison.Ordinal) == true;

    /// <summary>Builds auth detail from the chained messages, promoting any
    /// AADSTS code to the front (MSAL hides it behind generic sentences).</summary>
    public static string AuthDetail(Exception ex)
    {
        var lines = new List<string>();
        for (Exception? current = ex; current is not null; current = current.InnerException)
        {
            string first = FirstLine(current.Message);
            if (!string.IsNullOrWhiteSpace(first) && !lines.Contains(first))
            {
                lines.Add(first);
            }
        }

        string chain = string.Join(" | ", lines);
        var code = ExtractAadStsCode(chain);
        string detail = code is null ? chain : $"{code}: {chain}";
        return string.IsNullOrWhiteSpace(detail) ? "Authentication failed." : Truncate(detail, 300)!;
    }

    public static string? ExtractAadStsCode(string? text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return null;
        }

        var match = AadStsRegex().Match(text);
        return match.Success ? match.Groups[1].Value : null;
    }

    private static string FirstLine(string s)
    {
        int cut = s.IndexOfAny(['\r', '\n']);
        return cut < 0 ? s.Trim() : s[..cut].Trim();
    }

    private static string? Truncate(string? s, int max) =>
        s is null || s.Length <= max ? s : s[..max] + "…[truncated]";

    [GeneratedRegex("\"code\"\\s*:\\s*\"([^\"]+)\"", RegexOptions.IgnoreCase)]
    private static partial Regex GraphCodeRegex();

    [GeneratedRegex("(AADSTS\\d+)", RegexOptions.IgnoreCase)]
    private static partial Regex AadStsRegex();
}
