using Azure.Core;
using Azure.Identity;

namespace MicrosoftMcp.Common;

internal sealed class TokenCacheCredential(
    TokenCredential persistent,
    TokenCredential? memory,
    Action<string> warningSink,
    Func<CancellationToken, Task<AuthenticationRecord>>? authenticateAsync = null,
    Action<AuthenticationRecord>? authenticationRecordSink = null) : TokenCredential
{
    private int _useMemory;
    private int _warningWritten;

    public override AccessToken GetToken(
        TokenRequestContext requestContext,
        CancellationToken cancellationToken)
    {
        if (Volatile.Read(ref _useMemory) != 0)
        {
            return GetMemoryToken(requestContext, cancellationToken);
        }

        try
        {
            return persistent.GetToken(requestContext, cancellationToken);
        }
        catch (AuthenticationRequiredException) when (authenticateAsync is not null)
        {
            try
            {
                AuthenticationRecord record = authenticateAsync(cancellationToken).GetAwaiter().GetResult();
                authenticationRecordSink?.Invoke(record);
                return persistent.GetToken(requestContext, cancellationToken);
            }
            catch (Exception ex) when (TokenCredentialFactory.IsTokenCachePersistenceFailure(ex))
            {
                return SwitchToMemory(requestContext, cancellationToken, ex);
            }
        }
        catch (Exception ex) when (TokenCredentialFactory.IsTokenCachePersistenceFailure(ex))
        {
            return SwitchToMemory(requestContext, cancellationToken, ex);
        }
    }

    public override async ValueTask<AccessToken> GetTokenAsync(
        TokenRequestContext requestContext,
        CancellationToken cancellationToken)
    {
        if (Volatile.Read(ref _useMemory) != 0)
        {
            return await GetMemoryTokenAsync(requestContext, cancellationToken);
        }

        try
        {
            return await persistent.GetTokenAsync(requestContext, cancellationToken);
        }
        catch (AuthenticationRequiredException) when (authenticateAsync is not null)
        {
            try
            {
                AuthenticationRecord record = await authenticateAsync(cancellationToken);
                authenticationRecordSink?.Invoke(record);
                return await persistent.GetTokenAsync(requestContext, cancellationToken);
            }
            catch (Exception ex) when (TokenCredentialFactory.IsTokenCachePersistenceFailure(ex))
            {
                return await SwitchToMemoryAsync(requestContext, cancellationToken, ex);
            }
        }
        catch (Exception ex) when (TokenCredentialFactory.IsTokenCachePersistenceFailure(ex))
        {
            return await SwitchToMemoryAsync(requestContext, cancellationToken, ex);
        }
    }

    private AccessToken SwitchToMemory(
        TokenRequestContext requestContext,
        CancellationToken cancellationToken,
        Exception failure)
    {
        EnsureMemoryFallback(failure);
        Interlocked.Exchange(ref _useMemory, 1);
        WarnMemoryFallback();
        return memory!.GetToken(requestContext, cancellationToken);
    }

    private async ValueTask<AccessToken> SwitchToMemoryAsync(
        TokenRequestContext requestContext,
        CancellationToken cancellationToken,
        Exception failure)
    {
        EnsureMemoryFallback(failure);
        Interlocked.Exchange(ref _useMemory, 1);
        WarnMemoryFallback();
        return await memory!.GetTokenAsync(requestContext, cancellationToken);
    }

    private AccessToken GetMemoryToken(
        TokenRequestContext requestContext,
        CancellationToken cancellationToken)
    {
        EnsureMemoryFallback();
        return memory!.GetToken(requestContext, cancellationToken);
    }

    private async ValueTask<AccessToken> GetMemoryTokenAsync(
        TokenRequestContext requestContext,
        CancellationToken cancellationToken)
    {
        EnsureMemoryFallback();
        return await memory!.GetTokenAsync(requestContext, cancellationToken);
    }

    private void EnsureMemoryFallback(Exception? failure = null)
    {
        if (memory is null)
        {
            throw MailServiceException.AuthCacheUnavailable(failure);
        }
    }

    private void WarnMemoryFallback()
    {
        if (Interlocked.Exchange(ref _warningWritten, 1) == 0)
        {
            warningSink("[auth] Persistent token cache unavailable; using an in-memory cache. Tokens will not survive a process restart. Install a Secret Service or set Graph__UnsafeAllowUnencryptedTokenCache=true to persist an unencrypted cache.");
        }
    }
}