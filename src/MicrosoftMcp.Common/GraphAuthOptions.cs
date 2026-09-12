namespace MicrosoftMcp.Common;

public enum AuthMode
{
    Delegated,
    AppOnly
}

public enum DelegatedFlow
{
    Auto,
    InteractiveBrowser,
    DeviceCode
}

public enum AppCredentialKind
{
    ClientSecret,
    Certificate,
    ManagedIdentity
}

/// <summary>Binds to the "Graph" configuration section. Env override via Graph__*.</summary>
public sealed class GraphAuthOptions
{
    public const string SectionName = "Graph";

    public AuthMode AuthMode { get; set; } = AuthMode.Delegated;

    public string TenantId { get; set; } = string.Empty;

    public string ClientId { get; set; } = string.Empty;

    /// <summary>"me" for delegated self, or a UPN/user-id. Required for AppOnly.</summary>
    public string UserIdOrUpn { get; set; } = "me";

    public DelegatedFlow DelegatedFlow { get; set; } = DelegatedFlow.Auto;

    public string[] DelegatedScopes { get; set; } = ["Mail.Read", "Mail.ReadWrite"];

    public bool EnableTokenCache { get; set; } = true;

    public AppCredentialKind AppCredential { get; set; } = AppCredentialKind.ClientSecret;

    /// <summary>Via Graph__ClientSecret env var or user-secrets. Never commit.</summary>
    public string? ClientSecret { get; set; }
}
