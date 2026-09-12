using System.Text.RegularExpressions;
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
                api.ResponseStatusCode, ExtractGraphCode(api.Message), Truncate(api.Message, 300), operation),
            HttpRequestException http => MailServiceException.ServiceUnavailable(
                Truncate(http.Message, 200), http),
            TimeoutException timeout => MailServiceException.ServiceUnavailable(
                Truncate(timeout.Message, 200), timeout),
            ArgumentException arg => MailServiceException.InvalidRequest(
                Truncate(arg.Message, 300) ?? "Invalid argument.",
                "check ids, folder names and enum values (low|normal|high), then retry"),
            { } other when IsAuthNamespace(other) => MailServiceException.AuthFailed(
                Truncate(other.Message, 300) ?? "Authentication failed.", other),
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

    private static string? Truncate(string? s, int max) =>
        s is null || s.Length <= max ? s : s[..max] + "…[truncated]";

    [GeneratedRegex("\"code\"\\s*:\\s*\"([^\"]+)\"", RegexOptions.IgnoreCase)]
    private static partial Regex GraphCodeRegex();
}
