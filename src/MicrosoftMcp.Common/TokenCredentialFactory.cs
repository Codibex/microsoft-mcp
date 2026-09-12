using Azure.Core;
using Azure.Identity;

namespace MicrosoftMcp.Common;

public sealed class TokenCredentialFactory : ITokenCredentialProvider
{
    public TokenCredential GetCredential(GraphAuthOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);

        return options.AuthMode switch
        {
            AuthMode.Delegated => CreateDelegated(options),
            AuthMode.AppOnly => CreateAppOnly(options),
            _ => throw new InvalidOperationException($"Unsupported AuthMode '{options.AuthMode}'.")
        };
    }

    private static TokenCredential CreateDelegated(GraphAuthOptions options)
    {
        if (string.IsNullOrWhiteSpace(options.TenantId))
        {
            throw MailServiceException.AuthMisconfigured("Graph:TenantId is required for delegated auth.");
        }

        if (string.IsNullOrWhiteSpace(options.ClientId))
        {
            throw MailServiceException.AuthMisconfigured("Graph:ClientId is required for delegated auth.");
        }

        TokenCachePersistenceOptions? cache = options.EnableTokenCache
            ? new TokenCachePersistenceOptions { Name = "microsoft-mcp-outlook" }
            : null;

        if (options.DelegatedFlow == DelegatedFlow.DeviceCode)
        {
            return new DeviceCodeCredential(new DeviceCodeCredentialOptions
            {
                TenantId = options.TenantId,
                ClientId = options.ClientId,
                DeviceCodeCallback = (info, _) =>
                {
                    Console.Error.WriteLine($"[auth] {info.Message}");
                    return Task.CompletedTask;
                },
                TokenCachePersistenceOptions = cache
            });
        }

        // Auto and InteractiveBrowser behave the same locally: browser login.
        // DeviceCode is the explicit fallback for headless machines.
        return new InteractiveBrowserCredential(new InteractiveBrowserCredentialOptions
        {
            TenantId = options.TenantId,
            ClientId = options.ClientId,
            TokenCachePersistenceOptions = cache
        });
    }

    private static TokenCredential CreateAppOnly(GraphAuthOptions options)
    {
        if (string.IsNullOrWhiteSpace(options.TenantId))
        {
            throw MailServiceException.AuthMisconfigured("Graph:TenantId is required for app-only auth.");
        }

        if (string.IsNullOrWhiteSpace(options.ClientId))
        {
            throw MailServiceException.AuthMisconfigured("Graph:ClientId is required for app-only auth.");
        }

        return options.AppCredential switch
        {
            AppCredentialKind.ClientSecret => string.IsNullOrWhiteSpace(options.ClientSecret)
                ? throw MailServiceException.AuthMisconfigured("Graph:ClientSecret is required (use Graph__ClientSecret env var or user-secrets).")
                : new ClientSecretCredential(options.TenantId, options.ClientId, options.ClientSecret),
            AppCredentialKind.ManagedIdentity => new ManagedIdentityCredential(options.ClientId),
            AppCredentialKind.Certificate => throw new NotSupportedException(
                "Certificate auth is not wired yet. Use ClientSecret or ManagedIdentity."),
            _ => throw new InvalidOperationException($"Unsupported AppCredential '{options.AppCredential}'.")
        };
    }
}
