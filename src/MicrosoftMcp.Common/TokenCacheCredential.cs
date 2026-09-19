using Azure.Core;
using Azure.Identity;

namespace MicrosoftMcp.Common;

internal sealed class TokenCacheCredential(
    TokenCredential persistent,
    TokenCredential? memory,
    Action<string> warningSink,
    Func<TokenRequestContext, CancellationToken, Task<AuthenticationRecord>>? authenticateAsync = null,
    Action<AuthenticationRecord>? authenticationRecordSink = null,
    TimeSpan? persistentOperationTimeout = null) : TokenCredential
{
    private int _useMemory;
    private int _warningWritten;
    private readonly TimeSpan _persistentOperationTimeout = persistentOperationTimeout ?? TimeSpan.FromSeconds(10);

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
            return ExecutePersistent(
                token => persistent.GetToken(requestContext, token),
                cancellationToken);
        }
        catch (AuthenticationRequiredException) when (authenticateAsync is not null)
        {
            try
            {
                AuthenticationRecord record = authenticateAsync(requestContext, cancellationToken).GetAwaiter().GetResult();
                authenticationRecordSink?.Invoke(record);
                return ExecutePersistent(
                    token => persistent.GetToken(requestContext, token),
                    cancellationToken);
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
            return await ExecutePersistentAsync(
                token => persistent.GetTokenAsync(requestContext, token).AsTask(),
                cancellationToken);
        }
        catch (AuthenticationRequiredException) when (authenticateAsync is not null)
        {
            try
            {
                AuthenticationRecord record = await authenticateAsync(requestContext, cancellationToken);
                authenticationRecordSink?.Invoke(record);
                return await ExecutePersistentAsync(
                    token => persistent.GetTokenAsync(requestContext, token).AsTask(),
                    cancellationToken);
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
        WarnMemoryFallback(failure);
        return memory!.GetToken(requestContext, cancellationToken);
    }

    private async ValueTask<AccessToken> SwitchToMemoryAsync(
        TokenRequestContext requestContext,
        CancellationToken cancellationToken,
        Exception failure)
    {
        EnsureMemoryFallback(failure);
        Interlocked.Exchange(ref _useMemory, 1);
        WarnMemoryFallback(failure);
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
            throw GraphServiceException.AuthCacheUnavailable(failure);
        }
    }

    private T ExecutePersistent<T>(
        Func<CancellationToken, T> operation,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        using CancellationTokenSource timeoutSource =
            CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutSource.CancelAfter(_persistentOperationTimeout);

        Task<T> operationTask = Task.Run(
            () => operation(timeoutSource.Token),
            CancellationToken.None);

        try
        {
            return operationTask.WaitAsync(_persistentOperationTimeout, cancellationToken)
                .GetAwaiter().GetResult();
        }
        catch (TimeoutException) when (!cancellationToken.IsCancellationRequested)
        {
            throw CreateTimeoutException();
        }
        catch (OperationCanceledException) when (
            !cancellationToken.IsCancellationRequested && timeoutSource.IsCancellationRequested)
        {
            throw CreateTimeoutException();
        }
    }

    private async Task<T> ExecutePersistentAsync<T>(
        Func<CancellationToken, Task<T>> operation,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        using CancellationTokenSource timeoutSource =
            CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutSource.CancelAfter(_persistentOperationTimeout);

        Task<T> operationTask = Task.Run(
            () => operation(timeoutSource.Token),
            CancellationToken.None);

        try
        {
            return await operationTask.WaitAsync(_persistentOperationTimeout, cancellationToken)
                .ConfigureAwait(false);
        }
        catch (TimeoutException) when (!cancellationToken.IsCancellationRequested)
        {
            throw CreateTimeoutException();
        }
        catch (OperationCanceledException) when (
            !cancellationToken.IsCancellationRequested && timeoutSource.IsCancellationRequested)
        {
            throw CreateTimeoutException();
        }
    }

    private TimeoutException CreateTimeoutException() =>
        new($"Persistent token cache operation timed out after {_persistentOperationTimeout.TotalSeconds:0.###} seconds.");

    private void WarnMemoryFallback(Exception? failure = null)
    {
        if (Interlocked.Exchange(ref _warningWritten, 1) == 0)
        {
            string reason = failure is TimeoutException
                ? $" The persistent cache did not respond within {_persistentOperationTimeout.TotalSeconds:0.###} seconds."
                : string.Empty;
            warningSink($"[auth] Persistent token cache unavailable; using an in-memory cache.{reason} Tokens will not survive a process restart. Install a Secret Service or set Graph__UnsafeAllowUnencryptedTokenCache=true to persist an unencrypted cache.");
        }
    }
}