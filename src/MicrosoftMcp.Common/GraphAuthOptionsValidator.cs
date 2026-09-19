using Microsoft.Extensions.Options;

namespace MicrosoftMcp.Common;

public sealed class GraphAuthOptionsValidator : IValidateOptions<GraphAuthOptions>
{
    public ValidateOptionsResult Validate(string? name, GraphAuthOptions options)
    {
        if (string.IsNullOrWhiteSpace(options.TenantId))
        {
            return ValidateOptionsResult.Fail(
                "Graph:TenantId is required. Next: set it in appsettings.json (next to the binary), user-secrets, or Graph__TenantId env var.");
        }

        if (string.IsNullOrWhiteSpace(options.ClientId))
        {
            return ValidateOptionsResult.Fail(
                "Graph:ClientId is required. Next: set it in appsettings.json (next to the binary), user-secrets, or Graph__ClientId env var.");
        }

        if (options.AuthMode == AuthMode.AppOnly
            && string.Equals(options.UserIdOrUpn, "me", StringComparison.OrdinalIgnoreCase))
        {
            return ValidateOptionsResult.Fail(
                "Graph:UserIdOrUpn must be a UPN/user-id for AppOnly (not 'me'). Next: set Graph__UserIdOrUpn to the mailbox owner.");
        }

        if (options.AuthMode == AuthMode.Delegated
            && options.EnableTokenCache
            && (options.TokenCacheTimeoutSeconds is < 1 or > 120))
        {
            return ValidateOptionsResult.Fail(
                "Graph:TokenCacheTimeoutSeconds must be between 1 and 120 when the delegated token cache is enabled. Next: set Graph__TokenCacheTimeoutSeconds to a bounded value.");
        }

        return ValidateOptionsResult.Success;
    }
}
