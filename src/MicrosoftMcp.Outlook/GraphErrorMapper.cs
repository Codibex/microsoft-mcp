using System.Text.RegularExpressions;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Kiota.Abstractions;
using MicrosoftMcp.Common;

namespace MicrosoftMcp.Outlook;

/// <summary>Translates Graph/Kiota/transport exceptions into agent-actionable errors.</summary>
internal static partial class GraphErrorMapper
{
    private static readonly Dictionary<string, Func<int, string?, string, MailServiceException>> GraphCodeOverrides =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["ErrorItemNotFound"] = static (_, _, op) =>
                MailServiceException.MessageNotFound("<unknown>", op),
            ["ErrorMessageNotFound"] = static (_, _, op) =>
                MailServiceException.MessageNotFound("<unknown>", op),
            ["ErrorFolderNotFound"] = static (_, _, _) =>
                MailServiceException.FolderNotFound("<unknown>"),
            ["ErrorAccessDenied"] = static (s, c, _) =>
                MailServiceException.AccessDenied(s, c, null),
            ["ErrorInvalidIdMalformed"] = static (_, _, _) =>
                MailServiceException.InvalidRequest(
                    "Malformed Graph id.",
                    "use the 'id' field from search_emails, not internetMessageId or webLink"),
            ["MailboxNotEnabledForRESTAPI"] = static (_, c, d) =>
                MailServiceException.MailboxUnavailable(c, d),
            ["ErrorMailboxNotAssociated"] = static (_, c, d) =>
                MailServiceException.MailboxUnavailable(c, d),
            ["ErrorMailboxStoreUnknown"] = static (_, c, d) =>
                MailServiceException.MailboxUnavailable(c, d),
            ["ErrorNonExistentMailbox"] = static (_, c, d) =>
                MailServiceException.MailboxUnavailable(c, d),
            ["ErrorQuotaExceeded"] = static (_, _, _) =>
                MailServiceException.InvalidRequest(
                    "Mailbox quota exceeded.",
                    "free mailbox space or narrow the operation before retrying"),
        };

    internal static MailServiceException ToMailServiceException(Exception ex, string operation) =>
        ex switch
        {
            MailServiceException already => already,
            ApiException api => FromStatus(
                api.ResponseStatusCode, ResolveGraphCode(api), Truncate(api.Message, 300), operation),
            HttpRequestException http => MailServiceException.ServiceUnavailable(
                Truncate(http.Message, 200), http),
            TimeoutException timeout => MailServiceException.ServiceUnavailable(
                Truncate(timeout.Message, 200), timeout),
            ArgumentException arg => MailServiceException.InvalidRequest(
                Truncate(arg.Message, 300) ?? "Invalid argument.",
                "check ids, folder names and enum values (low|normal|high), then retry"),
            { } other when IsAuthNamespace(other) => MailServiceException.AuthFailed(
                AuthDetail(other), other),
            _ => MailServiceException.GraphError(0, ex.GetType().Name, Truncate(ex.Message, 300))
        };

    internal static MailServiceException FromStatus(
        int status, string? graphCode, string? detail, string operation)
    {
        if (graphCode is not null && GraphCodeOverrides.TryGetValue(graphCode, out var map))
        {
            return map(status, graphCode, operation);
        }

        return status switch
        {
            400 => MailServiceException.InvalidRequest(
                $"Graph rejected the request (400{(graphCode is null ? string.Empty : $", {graphCode}")}). {detail ?? string.Empty}".Trim(),
                "check ids, folder names and parameters, then retry with a narrower request"),
            401 => MailServiceException.AuthFailed(
                $"Graph authentication failed (401{(graphCode is null ? string.Empty : $", {graphCode}")}). {detail ?? string.Empty}".Trim()),
            403 => MailServiceException.AccessDenied(status, graphCode, detail),
            404 => MailServiceException.MessageNotFound("<unknown>", operation),
            409 => MailServiceException.Conflict(operation, detail),
            429 => MailServiceException.Throttled(detail),
            >= 500 => MailServiceException.ServiceUnavailable(
                $"HTTP {status}{(graphCode is null ? string.Empty : $", {graphCode}")}. {detail ?? string.Empty}".Trim()),
            _ => MailServiceException.GraphError(status, graphCode, detail)
        };
    }

    /// <summary>Prefers the structured ODataError code; falls back to regex
    /// over the message text (Kiota puts the Graph body there).</summary>
    internal static string? ResolveGraphCode(ApiException api) =>
        api is ODataError odata && !string.IsNullOrWhiteSpace(odata.Error?.Code)
            ? odata.Error.Code
            : ExtractGraphCode(api.Message);

    internal static string? ExtractGraphCode(string? message)
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
    internal static string AuthDetail(Exception ex)
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

    internal static string? ExtractAadStsCode(string? text)
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
