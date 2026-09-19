using Azure.Core;
using Azure.Identity;

namespace MicrosoftMcp.Common;

public sealed class TokenCredentialFactory : ITokenCredentialProvider
{
    private readonly Action<string> _warningSink;
    private readonly Func<GraphAuthOptions, TokenCachePersistenceOptions?, AuthenticationRecord?, TokenCredential> _credentialFactory;
    private readonly AuthenticationRecordStore _authenticationRecordStore;
    private int _unsafeCacheWarningWritten;

    public TokenCredentialFactory() : this(Console.Error.WriteLine, null)
    {
    }

    internal TokenCredentialFactory(
        Action<string> warningSink,
        Func<GraphAuthOptions, TokenCachePersistenceOptions?, TokenCredential>? credentialFactory = null,
        AuthenticationRecordStore? authenticationRecordStore = null)
    {
        _warningSink = warningSink ?? throw new ArgumentNullException(nameof(warningSink));
        _credentialFactory = credentialFactory is null
            ? CreateDelegatedCredential
            : (options, cache, _) => credentialFactory(options, cache);
        _authenticationRecordStore = authenticationRecordStore ?? new AuthenticationRecordStore();
    }

    internal TokenCredentialFactory(
        Action<string> warningSink,
        Func<GraphAuthOptions, TokenCachePersistenceOptions?, AuthenticationRecord?, TokenCredential> credentialFactory,
        AuthenticationRecordStore authenticationRecordStore)
    {
        _warningSink = warningSink ?? throw new ArgumentNullException(nameof(warningSink));
        _credentialFactory = credentialFactory ?? throw new ArgumentNullException(nameof(credentialFactory));
        _authenticationRecordStore = authenticationRecordStore ?? throw new ArgumentNullException(nameof(authenticationRecordStore));
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
            throw GraphServiceException.AuthMisconfigured("Graph:TenantId is required for delegated auth.");
        }

        if (string.IsNullOrWhiteSpace(options.ClientId))
        {
            throw GraphServiceException.AuthMisconfigured("Graph:ClientId is required for delegated auth.");
        }

        if (!options.EnableTokenCache)
        {
            return _credentialFactory(options, null, null);
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

        AuthenticationRecord? authenticationRecord = _authenticationRecordStore.Load();
        if (authenticationRecord is not null
            && !IsAuthenticationRecordCompatible(authenticationRecord, options))
        {
            authenticationRecord = null;
        }

        try
        {
            TokenCredential persistent = _credentialFactory(options, cache, authenticationRecord);
            TokenCredential? memory = options.FallbackToMemoryTokenCache
                ? _credentialFactory(options, null, null)
                : null;
            return new TokenCacheCredential(
                persistent,
                memory,
                _warningSink,
                GetAuthenticationHandler(persistent),
                record => _authenticationRecordStore.Save(record, _warningSink),
                TimeSpan.FromSeconds(options.TokenCacheTimeoutSeconds));
        }
        catch (Exception ex) when (IsTokenCachePersistenceFailure(ex))
        {
            if (options.FallbackToMemoryTokenCache)
            {
                WarnMemoryCacheFallback();
                return _credentialFactory(options, null, null);
            }

            throw GraphServiceException.AuthCacheUnavailable(ex);
        }
    }

    private static TokenCredential CreateDelegatedCredential(
        GraphAuthOptions options,
        TokenCachePersistenceOptions? cache,
        AuthenticationRecord? authenticationRecord)
    {
        if (options.DelegatedFlow == DelegatedFlow.DeviceCode)
        {
            return new DeviceCodeCredential(new DeviceCodeCredentialOptions
            {
                TenantId = options.TenantId,
                ClientId = options.ClientId,
                AuthenticationRecord = authenticationRecord,
                DisableAutomaticAuthentication = cache is not null,
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
            AuthenticationRecord = authenticationRecord,
            DisableAutomaticAuthentication = cache is not null,
            TokenCachePersistenceOptions = cache
        });
    }

    private static Func<TokenRequestContext, CancellationToken, Task<AuthenticationRecord>>? GetAuthenticationHandler(TokenCredential credential) =>
        credential switch
        {
            DeviceCodeCredential deviceCode => deviceCode.AuthenticateAsync,
            InteractiveBrowserCredential browser => browser.AuthenticateAsync,
            _ => null
        };

    private static bool IsAuthenticationRecordCompatible(
        AuthenticationRecord record,
        GraphAuthOptions options)
    {
        if (!string.Equals(record.ClientId, options.ClientId, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return options.TenantId.Equals("common", StringComparison.OrdinalIgnoreCase)
            || options.TenantId.Equals("consumers", StringComparison.OrdinalIgnoreCase)
            || options.TenantId.Equals("organizations", StringComparison.OrdinalIgnoreCase)
            || string.Equals(record.TenantId, options.TenantId, StringComparison.OrdinalIgnoreCase);
    }

    internal static bool IsTokenCachePersistenceFailure(Exception ex)
    {
        for (Exception? current = ex; current is not null; current = current.InnerException)
        {
            string message = current.Message;
            if (current is TimeoutException
                || message.Contains("Persistence check failed", StringComparison.OrdinalIgnoreCase)
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
            throw GraphServiceException.AuthMisconfigured("Graph:TenantId is required for app-only auth.");
        }

        if (string.IsNullOrWhiteSpace(options.ClientId))
        {
            throw GraphServiceException.AuthMisconfigured("Graph:ClientId is required for app-only auth.");
        }

        return options.AppCredential switch
        {
            AppCredentialKind.ClientSecret => string.IsNullOrWhiteSpace(options.ClientSecret)
                ? throw GraphServiceException.AuthMisconfigured("Graph:ClientSecret is required (use Graph__ClientSecret env var or user-secrets).")
                : new ClientSecretCredential(options.TenantId, options.ClientId, options.ClientSecret),
            AppCredentialKind.ManagedIdentity => new ManagedIdentityCredential(options.ClientId),
            AppCredentialKind.Certificate => throw new NotSupportedException(
                "Certificate auth is not wired yet. Use ClientSecret or ManagedIdentity."),
            _ => throw new InvalidOperationException($"Unsupported AppCredential '{options.AppCredential}'.")
        };
    }
}
