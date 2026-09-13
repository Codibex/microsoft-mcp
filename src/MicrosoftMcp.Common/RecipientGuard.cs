namespace MicrosoftMcp.Common;

/// <summary>Server-side mail recipient check. Runs in code (not via prompt), so the LLM
/// cannot talk its way around it. Pure and unit-tested. Teams send will reuse the
/// same policy via tenant-member checks (see docs/teams.md).</summary>
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
            if (!IsAllowed(domain, policy.AllowedRecipientDomains))
            {
                throw MailServiceException.InvalidRequest(
                    $"Recipient '{recipient}' is outside the allowed domains ({string.Join(", ", policy.AllowedRecipientDomains)}).",
                    "use an internal recipient address, or ask your admin to extend policy.json");
            }
        }
    }

    public static string DomainOf(string address)
    {
        int at = address.LastIndexOf('@');
        if (at <= 0 || at == address.Length - 1)
        {
            throw MailServiceException.InvalidRequest(
                $"Recipient '{address}' is not a valid email address.",
                "pass addresses like \"name@firma.de\"");
        }

        string domain = address[(at + 1)..].Trim().ToLowerInvariant();
        if (domain.Length == 0 || domain.Contains(' ') || domain.Contains('@') || !domain.Contains('.'))
        {
            throw MailServiceException.InvalidRequest(
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
}
