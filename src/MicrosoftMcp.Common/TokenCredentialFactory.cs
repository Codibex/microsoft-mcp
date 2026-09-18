using Azure.Core;
using Azure.Identity;

namespace MicrosoftMcp.Common;

public sealed class TokenCredentialFactory : ITokenCredentialProvider
{
    private readonly Action<string> _warningSink;
    private int _unsafeCacheWarningWritten;

    public TokenCredentialFactory() : this(Console.Error.WriteLine)
    {
    }

    internal TokenCredentialFactory(Action<string> warningSink)
    {
        _warningSink = warningSink ?? throw new ArgumentNullException(nameof(warningSink));
    }

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

    private TokenCredential CreateDelegated(GraphAuthOptions options)
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

        if (options.UnsafeAllowUnencryptedTokenCache)
        {
            WarnUnencryptedTokenCache();
        }

        TokenCachePersistenceOptions cache = new()
        {
            Name = "microsoft-mcp-graph",
            UnsafeAllowUnencryptedStorage = options.UnsafeAllowUnencryptedTokenCache
        };

        try
        {
            TokenCredential persistent = CreateDelegatedCredential(options, cache);
            TokenCredential? memory = options.FallbackToMemoryTokenCache
                ? CreateDelegatedCredential(options, cache: null)
                : null;
            return new TokenCacheCredential(persistent, memory, _warningSink);
        }
        catch (Exception ex) when (IsTokenCachePersistenceFailure(ex))
        {
            if (options.FallbackToMemoryTokenCache)
            {
                WarnMemoryCacheFallback();
                return CreateDelegatedCredential(options, cache: null);
            }

            throw MailServiceException.AuthCacheUnavailable(ex);
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

    internal static bool IsTokenCachePersistenceFailure(Exception ex)
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

    private void WarnMemoryCacheFallback() =>
        _warningSink("[auth] Persistent token cache unavailable; using an in-memory cache. Tokens will not survive a process restart. Install a Secret Service or set Graph__UnsafeAllowUnencryptedTokenCache=true to persist an unencrypted cache.");

    private void WarnUnencryptedTokenCache()
    {
        if (Interlocked.Exchange(ref _unsafeCacheWarningWritten, 1) == 0)
        {
            _warningSink("[auth] WARNING: unencrypted token-cache storage is enabled. Protect the cache file and prefer the encrypted OS cache or the in-memory fallback.");
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
