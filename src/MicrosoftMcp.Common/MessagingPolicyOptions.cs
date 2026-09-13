namespace MicrosoftMcp.Common;

/// <summary>Admin-owned enterprise policy shared by all messaging servers
/// (Outlook drafts today, Teams send in the future): internal recipients +
/// AI disclosure.
///
/// Binds ONLY from the protected <c>policy.json</c> file (see <see cref="PolicyFile"/>).
/// Env vars / user-secrets are deliberately NOT consulted for these settings,
/// because <c>mcp.json</c> is user-writable and an LLM with file access could
/// otherwise override the admin policy.</summary>
public sealed class MessagingPolicyOptions
{
    /// <summary>When true, messages may only target internal recipients
    /// (mail: <see cref="AllowedRecipientDomains"/>; Teams: same-tenant members).</summary>
    public bool RequireInternalRecipients { get; set; }

    /// <summary>Allowed mail recipient domains, e.g. ["firma.de"]. Subdomains match
    /// (mail.firma.de is covered by firma.de). Compared case-insensitively.</summary>
    public string[] AllowedRecipientDomains { get; set; } = [];

    /// <summary>When true, <see cref="AiDisclosureText"/> is appended to every
    /// created message by the server (not by the LLM, so it cannot be omitted).</summary>
    public bool AiDisclosureEnabled { get; set; }

    /// <summary>Admin-controlled disclosure text. Not secret, but integrity-protected
    /// via the file ownership check in <see cref="PolicyFile"/>.</summary>
    public string AiDisclosureText { get; set; } = string.Empty;

    /// <summary>Whether this policy restricts anything (only then the file
    /// protection check applies).</summary>
    public bool IsRestrictive => RequireInternalRecipients || AiDisclosureEnabled;
}
