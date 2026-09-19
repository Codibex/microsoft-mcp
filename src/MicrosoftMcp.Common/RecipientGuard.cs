namespace MicrosoftMcp.Common;

/// <summary>Server-side recipient check. Runs in code (not via prompt), so the
/// LLM cannot talk its way around it. Pure and unit-tested. Teams send and
/// calendar attendees reuse the same policy.</summary>
public static class RecipientGuard
{
    public static void ValidateRecipients(IReadOnlyList<string> recipients, MessagingPolicyOptions policy)
    {
        if (!policy.RequireInternalRecipients)
        {
            return;
        }

        foreach (string recipient in recipients)
        {
            string domain = DomainOf(recipient);
            if (!IsAllowedRecipient(recipient, domain, policy))
            {
                throw GraphServiceException.InvalidRequest(
                    $"Recipient '{recipient}' is outside the allowed domains and exact addresses configured by policy.json.",
                    "use an internal recipient address, or ask your admin to extend policy.json");
            }
        }
    }

    public static string DomainOf(string address)
    {
        // Exactly one '@': LastIndexOf would let 'a@external@firma.de' pass as internal.
        if (address.Count(c => c == '@') != 1)
        {
            throw GraphServiceException.InvalidRequest(
                $"Recipient '{address}' is not a valid email address.",
                "pass addresses like \"name@firma.de\"");
        }

        int at = address.IndexOf('@');
        if (at <= 0 || at == address.Length - 1)
        {
            throw GraphServiceException.InvalidRequest(
                $"Recipient '{address}' is not a valid email address.",
                "pass addresses like \"name@firma.de\"");
        }

        string domain = address[(at + 1)..].Trim().ToLowerInvariant();
        if (domain.Length == 0 || domain.Contains(' ') || domain.Contains('@') || !domain.Contains('.'))
        {
            throw GraphServiceException.InvalidRequest(
                $"Recipient '{address}' is not a valid email address.",
                "pass addresses like \"name@firma.de\"");
        }

        return domain;
    }

    public static bool IsAllowed(string domain, string[] allowed) =>
        allowed.Any(a =>
        {
            string normalized = a.Trim().TrimStart('.').ToLowerInvariant();
            return normalized.Length > 0 &&
                (domain == normalized || domain.EndsWith("." + normalized, StringComparison.Ordinal));
        });

    private static bool IsAllowedRecipient(
        string address, string domain, MessagingPolicyOptions policy) =>
        policy.AllowedRecipientAddresses.Any(allowed =>
            string.Equals(allowed.Trim(), address.Trim(), StringComparison.OrdinalIgnoreCase))
        || IsAllowed(domain, policy.AllowedRecipientDomains);
}
