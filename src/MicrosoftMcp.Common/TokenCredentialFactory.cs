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

        if (!options.EnableTokenCache)
        {
            return CreateDelegatedCredential(options, cache: null);
        }

        TokenCachePersistenceOptions cache = new()
        {
            Name = "microsoft-mcp-graph",
            UnsafeAllowUnencryptedStorage = options.UnsafeAllowUnencryptedTokenCache
        };

        try
        {
            TokenCredential persistent = CreateDelegatedCredential(options, cache);
            if (!options.FallbackToMemoryTokenCache)
            {
                return persistent;
            }

            return new CacheFallbackCredential(
                persistent,
                CreateDelegatedCredential(options, cache: null));
        }
        catch (Exception ex) when (options.FallbackToMemoryTokenCache && IsTokenCachePersistenceFailure(ex))
        {
            WarnMemoryCacheFallback();
            return CreateDelegatedCredential(options, cache: null);
        }
    }

    private static TokenCredential CreateDelegatedCredential(
        GraphAuthOptions options,
        TokenCachePersistenceOptions? cache)
    {
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

    private static bool IsTokenCachePersistenceFailure(Exception ex)
    {
        for (Exception? current = ex; current is not null; current = current.InnerException)
        {
            string message = current.Message;
            if (message.Contains("Persistence check failed", StringComparison.OrdinalIgnoreCase)
                || message.Contains("libsecret", StringComparison.OrdinalIgnoreCase)
                || message.Contains("Secret Service", StringComparison.OrdinalIgnoreCase)
                || message.Contains("keyring", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static void WarnMemoryCacheFallback() =>
        Console.Error.WriteLine("[auth] Persistent token cache unavailable; using an in-memory cache. Tokens will not survive a process restart. Install a Secret Service or set Graph__UnsafeAllowUnencryptedTokenCache=true to persist an unencrypted cache.");

    private sealed class CacheFallbackCredential(
        TokenCredential persistent,
        TokenCredential memory) : TokenCredential
    {
        private int _useMemory;
        private int _warningWritten;

        public override AccessToken GetToken(
            TokenRequestContext requestContext,
            CancellationToken cancellationToken)
        {
            if (Volatile.Read(ref _useMemory) != 0)
            {
                return memory.GetToken(requestContext, cancellationToken);
            }

            try
            {
                return persistent.GetToken(requestContext, cancellationToken);
            }
            catch (Exception ex) when (IsTokenCachePersistenceFailure(ex))
            {
                SwitchToMemory();
                return memory.GetToken(requestContext, cancellationToken);
            }
        }

        public override async ValueTask<AccessToken> GetTokenAsync(
            TokenRequestContext requestContext,
            CancellationToken cancellationToken)
        {
            if (Volatile.Read(ref _useMemory) != 0)
            {
                return await memory.GetTokenAsync(requestContext, cancellationToken);
            }

            try
            {
                return await persistent.GetTokenAsync(requestContext, cancellationToken);
            }
            catch (Exception ex) when (IsTokenCachePersistenceFailure(ex))
            {
                SwitchToMemory();
                return await memory.GetTokenAsync(requestContext, cancellationToken);
            }
        }

        private void SwitchToMemory()
        {
            Interlocked.Exchange(ref _useMemory, 1);
            if (Interlocked.Exchange(ref _warningWritten, 1) == 0)
            {
                WarnMemoryCacheFallback();
            }
        }
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
