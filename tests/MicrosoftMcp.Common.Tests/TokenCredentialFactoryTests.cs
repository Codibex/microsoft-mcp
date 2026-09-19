using Azure.Core;
using Azure.Identity;
using AwesomeAssertions;
using MicrosoftMcp.Common;
using NSubstitute;
using System.IO.Abstractions;

namespace MicrosoftMcp.Common.Tests;

public sealed class TokenCredentialFactoryTests
{
    private readonly TokenCredentialFactory _sut = new();

    [Fact]
    public void Null_options_throws()
    {
        Action act = () => _sut.GetCredential(null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void Delegated_without_tenant_throws()
    {
        var options = new GraphAuthOptions { AuthMode = AuthMode.Delegated, ClientId = "c" };
        Action act = () => _sut.GetCredential(options);
        act.Should().Throw<GraphServiceException>().Where(e => e.Code == "auth-misconfigured").WithMessage("*TenantId*");
    }

    [Fact]
    public void Delegated_without_client_throws()
    {
        var options = new GraphAuthOptions { AuthMode = AuthMode.Delegated, TenantId = "t" };
        Action act = () => _sut.GetCredential(options);
        act.Should().Throw<GraphServiceException>().Where(e => e.Code == "auth-misconfigured").WithMessage("*ClientId*");
    }

    [Fact]
    public void Delegated_devicecode_returns_credential_without_network()
    {
        var options = new GraphAuthOptions
        {
            AuthMode = AuthMode.Delegated,
            DelegatedFlow = DelegatedFlow.DeviceCode,
            TenantId = "t",
            ClientId = "c"
        };

        _sut.GetCredential(options).Should().NotBeNull();
    }

    [Fact]
    public void AppOnly_without_secret_throws()
    {
        var options = new GraphAuthOptions
        {
            AuthMode = AuthMode.AppOnly,
            TenantId = "t",
            ClientId = "c",
            UserIdOrUpn = "user@tenant"
        };

        Action act = () => _sut.GetCredential(options);
        act.Should().Throw<GraphServiceException>().Where(e => e.Code == "auth-misconfigured").WithMessage("*ClientSecret*");
    }

    [Fact]
    public void AppOnly_certificate_is_not_supported_yet()
    {
        var options = new GraphAuthOptions
        {
            AuthMode = AuthMode.AppOnly,
            AppCredential = AppCredentialKind.Certificate,
            TenantId = "t",
            ClientId = "c",
            UserIdOrUpn = "user@tenant"
        };

        Action act = () => _sut.GetCredential(options);
        act.Should().Throw<NotSupportedException>();
    }

    [Fact]
    public void AppOnly_secret_returns_credential_without_network()
    {
        var options = new GraphAuthOptions
        {
            AuthMode = AuthMode.AppOnly,
            TenantId = "t",
            ClientId = "c",
            UserIdOrUpn = "user@tenant",
            ClientSecret = "s"
        };

        _sut.GetCredential(options).Should().NotBeNull();
    }

    [Fact]
    public void Token_cache_fallback_switches_to_memory_and_warns_once()
    {
        List<string> warnings = [];
        var persistent = new FailingCredential("Persistence check failed: libsecret");
        var memory = new CountingCredential();
        var credential = new TokenCacheCredential(persistent, memory, warnings.Add);
        var context = new TokenRequestContext(["scope"]);

        credential.GetToken(context, CancellationToken.None);
        credential.GetToken(context, CancellationToken.None);

        memory.Calls.Should().Be(2);
        warnings.Count.Should().Be(1);
        warnings[0].Should().Contain("in-memory cache");
    }

    [Fact]
    public async Task Token_cache_async_timeout_switches_to_memory()
    {
        List<string> warnings = [];
        var persistent = new HangingCredential();
        var memory = new CountingCredential();
        var credential = new TokenCacheCredential(
            persistent,
            memory,
            warnings.Add,
            persistentOperationTimeout: TimeSpan.FromMilliseconds(25));

        try
        {
            AccessToken token = await credential.GetTokenAsync(
                new TokenRequestContext(["scope"]),
                CancellationToken.None).AsTask().WaitAsync(TimeSpan.FromSeconds(1));

            token.Token.Should().Be("token");
            memory.Calls.Should().Be(1);
            warnings.Should().ContainSingle().Which.Should().Contain("did not respond within");

            AccessToken secondToken = await credential.GetTokenAsync(
                new TokenRequestContext(["scope"]),
                CancellationToken.None);

            secondToken.Token.Should().Be("token");
            persistent.Calls.Should().Be(1);
            memory.Calls.Should().Be(2);
        }
        finally
        {
            persistent.Release();
        }
    }

    [Fact]
    public async Task Token_cache_sync_timeout_switches_to_memory()
    {
        List<string> warnings = [];
        var persistent = new HangingCredential();
        var memory = new CountingCredential();
        var credential = new TokenCacheCredential(
            persistent,
            memory,
            warnings.Add,
            persistentOperationTimeout: TimeSpan.FromMilliseconds(25));

        try
        {
            AccessToken token = await Task.Run(() => credential.GetToken(
                new TokenRequestContext(["scope"]),
                CancellationToken.None)).WaitAsync(TimeSpan.FromSeconds(1));

            token.Token.Should().Be("token");
            memory.Calls.Should().Be(1);
            warnings.Should().ContainSingle().Which.Should().Contain("did not respond within");
        }
        finally
        {
            persistent.Release();
        }
    }

    [Fact]
    public async Task Token_cache_timeout_in_strict_mode_returns_actionable_error()
    {
        var persistent = new HangingCredential();
        var credential = new TokenCacheCredential(
            persistent,
            memory: null,
            _ => { },
            persistentOperationTimeout: TimeSpan.FromMilliseconds(25));

        try
        {
            GraphServiceException error = await Assert.ThrowsAsync<GraphServiceException>(() =>
                credential.GetTokenAsync(
                    new TokenRequestContext(["scope"]),
                    CancellationToken.None).AsTask().WaitAsync(TimeSpan.FromSeconds(1)));

            error.Code.Should().Be("auth-cache-unavailable");
            error.Message.Should().Contain("FallbackToMemoryTokenCache");
        }
        finally
        {
            persistent.Release();
        }
    }

    [Fact]
    public async Task Token_cache_timeout_does_not_replace_caller_cancellation()
    {
        List<string> warnings = [];
        var persistent = new HangingCredential();
        var memory = new CountingCredential();
        var credential = new TokenCacheCredential(
            persistent,
            memory,
            warnings.Add,
            persistentOperationTimeout: TimeSpan.FromSeconds(1));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => credential.GetTokenAsync(
            new TokenRequestContext(["scope"]),
            cancellation.Token).AsTask());

        memory.Calls.Should().Be(0);
        warnings.Should().BeEmpty();
        persistent.Release();
    }

    [Fact]
    public void Token_cache_construction_failure_falls_back_to_memory()
    {
        List<string> warnings = [];
        var factory = new TokenCredentialFactory(
            warnings.Add,
            (_, cache) => cache is null
                ? new CountingCredential()
                : throw new InvalidOperationException("Persistence check failed: libsecret"));
        var options = new GraphAuthOptions
        {
            AuthMode = AuthMode.Delegated,
            TenantId = "t",
            ClientId = "c"
        };

        TokenCredential credential = factory.GetCredential(options);
        credential.GetToken(new TokenRequestContext(["scope"]), CancellationToken.None).Token
            .Should().Be("token");

        warnings.Count.Should().Be(1);
        warnings[0].Should().Contain("in-memory cache");
    }

    [Fact]
    public void Unrelated_construction_timeout_is_not_treated_as_cache_failure()
    {
        List<string> warnings = [];
        var factory = new TokenCredentialFactory(
            warnings.Add,
            (_, cache) => cache is null
                ? new CountingCredential()
                : throw new TimeoutException("unrelated timeout"));
        var options = new GraphAuthOptions
        {
            AuthMode = AuthMode.Delegated,
            TenantId = "t",
            ClientId = "c"
        };

        Action act = () => factory.GetCredential(options);

        act.Should().Throw<TimeoutException>().WithMessage("unrelated timeout");
        warnings.Should().BeEmpty();
    }

    [Fact]
    public void Token_cache_strict_mode_throws_actionable_cache_error()
    {
        var credential = new TokenCacheCredential(
            new FailingCredential("Persistence check failed: libsecret"),
            memory: null,
            _ => { });

        Action act = () => credential.GetToken(new TokenRequestContext(["scope"]), CancellationToken.None);

        act.Should().Throw<GraphServiceException>()
            .Where(e => e.Code == "auth-cache-unavailable")
            .WithMessage("*FallbackToMemoryTokenCache*");
    }

    [Fact]
    public async Task Unrelated_token_timeout_is_not_treated_as_cache_failure()
    {
        List<string> warnings = [];
        var memory = new CountingCredential();
        var credential = new TokenCacheCredential(
            new TimeoutCredential(),
            memory,
            warnings.Add,
            persistentOperationTimeout: TimeSpan.FromSeconds(1));

        await Assert.ThrowsAsync<TimeoutException>(() => credential.GetTokenAsync(
            new TokenRequestContext(["scope"]),
            CancellationToken.None).AsTask());

        memory.Calls.Should().Be(0);
        warnings.Should().BeEmpty();
    }

    [Fact]
    public void Token_cache_authenticates_and_saves_record_when_silent_authentication_is_required()
    {
        string directory = Path.Combine(Path.GetTempPath(), $"microsoft-mcp-{Guid.NewGuid():N}");
        string path = Path.Combine(directory, "authentication-record.json");
        AuthenticationRecord record = CreateAuthenticationRecord();
        var persistent = new AuthenticationRequiredCredential();
        TokenRequestContext? observedContext = null;
        int authenticateCalls = 0;
        try
        {
            var store = new AuthenticationRecordStore(path);
            List<string> warnings = [];
            var credential = new TokenCacheCredential(
                persistent,
                memory: null,
                _ => { },
                (context, _) =>
                {
                    observedContext = context;
                    authenticateCalls++;
                    persistent.Authenticated = true;
                    return Task.FromResult(record);
                },
                authenticatedRecord => store.Save(authenticatedRecord, warnings.Add));

            AccessToken token = credential.GetToken(new TokenRequestContext(["scope"]), CancellationToken.None);

            token.Token.Should().Be("token");
            authenticateCalls.Should().Be(1);
            observedContext.Should().NotBeNull();
            observedContext!.Value.Scopes.Should().Contain("scope");
            warnings.Should().BeEmpty();

            AuthenticationRecord? restored = new AuthenticationRecordStore(path).Load();
            restored.Should().NotBeNull();
            restored!.ClientId.Should().Be(record.ClientId);
            restored.HomeAccountId.Should().Be(record.HomeAccountId);
        }
        finally
        {
            File.Delete(path);
            Directory.Delete(directory);
        }
    }

    [Fact]
    public void Token_cache_authentication_persistence_failure_falls_back_to_memory()
    {
        List<string> warnings = [];
        var persistent = new AuthenticationRequiredCredential();
        var memory = new CountingCredential();
        var credential = new TokenCacheCredential(
            persistent,
            memory,
            warnings.Add,
            (_, _) => throw new InvalidOperationException("Persistence check failed: libsecret"));

        AccessToken token = credential.GetToken(new TokenRequestContext(["scope"]), CancellationToken.None);

        token.Token.Should().Be("token");
        memory.Calls.Should().Be(1);
        warnings.Count.Should().Be(1);
    }

    [Fact]
    public async Task Token_cache_async_authenticates_and_saves_record_when_silent_authentication_is_required()
    {
        var persistent = new AuthenticationRequiredCredential();
        TokenRequestContext? observedContext = null;
        int authenticateCalls = 0;
        int savedRecords = 0;
        var credential = new TokenCacheCredential(
            persistent,
            memory: null,
            _ => { },
            (context, _) =>
            {
                observedContext = context;
                authenticateCalls++;
                persistent.Authenticated = true;
                return Task.FromResult(CreateAuthenticationRecord());
            },
            _ => savedRecords++);

        AccessToken token = await credential.GetTokenAsync(new TokenRequestContext(["scope"]), CancellationToken.None);

        token.Token.Should().Be("token");
        authenticateCalls.Should().Be(1);
        observedContext.Should().NotBeNull();
        observedContext!.Value.Scopes.Should().Contain("scope");
        savedRecords.Should().Be(1);
    }

    [Fact]
    public async Task Token_cache_async_authentication_persistence_failure_falls_back_to_memory()
    {
        List<string> warnings = [];
        var persistent = new AuthenticationRequiredCredential();
        var memory = new CountingCredential();
        var credential = new TokenCacheCredential(
            persistent,
            memory,
            warnings.Add,
            (_, _) => Task.FromException<AuthenticationRecord>(
                new InvalidOperationException("Persistence check failed: libsecret")));

        AccessToken token = await credential.GetTokenAsync(new TokenRequestContext(["scope"]), CancellationToken.None);

        token.Token.Should().Be("token");
        memory.Calls.Should().Be(1);
        warnings.Count.Should().Be(1);
    }

    [Fact]
    public void Token_cache_retry_persistence_failure_falls_back_to_memory()
    {
        List<string> warnings = [];
        var persistent = new AuthenticationRequiredCredential
        {
            FailureAfterAuthentication = "Persistence check failed: libsecret"
        };
        var memory = new CountingCredential();
        var credential = new TokenCacheCredential(
            persistent,
            memory,
            warnings.Add,
            (_, _) =>
            {
                persistent.Authenticated = true;
                return Task.FromResult(CreateAuthenticationRecord());
            });

        AccessToken token = credential.GetToken(new TokenRequestContext(["scope"]), CancellationToken.None);

        token.Token.Should().Be("token");
        memory.Calls.Should().Be(1);
        warnings.Count.Should().Be(1);
    }

    [Fact]
    public async Task Token_cache_async_retry_persistence_failure_falls_back_to_memory()
    {
        List<string> warnings = [];
        var persistent = new AuthenticationRequiredCredential
        {
            FailureAfterAuthentication = "Persistence check failed: libsecret"
        };
        var memory = new CountingCredential();
        var credential = new TokenCacheCredential(
            persistent,
            memory,
            warnings.Add,
            (_, _) =>
            {
                persistent.Authenticated = true;
                return Task.FromResult(CreateAuthenticationRecord());
            });

        AccessToken token = await credential.GetTokenAsync(new TokenRequestContext(["scope"]), CancellationToken.None);

        token.Token.Should().Be("token");
        memory.Calls.Should().Be(1);
        warnings.Count.Should().Be(1);
    }

    [Fact]
    public void New_factory_instance_restores_record_for_persistent_credential()
    {
        string directory = Path.Combine(Path.GetTempPath(), $"microsoft-mcp-{Guid.NewGuid():N}");
        string path = Path.Combine(directory, "authentication-record.json");
        AuthenticationRecord record = CreateAuthenticationRecord();
        AuthenticationRecord? observed = null;
        try
        {
            var store = new AuthenticationRecordStore(path);
            List<string> warnings = [];
            store.Save(record, warnings.Add);
            warnings.Should().BeEmpty();

            var factory = new TokenCredentialFactory(
                _ => { },
                (options, cache, authenticationRecord) =>
                {
                    observed = authenticationRecord;
                    cache.Should().NotBeNull();
                    return new CountingCredential();
                },
                new AuthenticationRecordStore(path));

            factory.GetCredential(new GraphAuthOptions
            {
                AuthMode = AuthMode.Delegated,
                TenantId = record.TenantId,
                ClientId = record.ClientId,
                FallbackToMemoryTokenCache = false
            });

            observed.Should().NotBeNull();
            observed!.Username.Should().Be(record.Username);
            observed.HomeAccountId.Should().Be(record.HomeAccountId);
        }
        finally
        {
            File.Delete(path);
            Directory.Delete(directory);
        }
    }

    [Fact]
    public void New_factory_instance_rejects_record_from_different_concrete_tenant()
    {
        string directory = Path.Combine(Path.GetTempPath(), $"microsoft-mcp-{Guid.NewGuid():N}");
        string path = Path.Combine(directory, "authentication-record.json");
        AuthenticationRecord record = CreateAuthenticationRecord();
        AuthenticationRecord? observed = null;
        try
        {
            new AuthenticationRecordStore(path).Save(record, _ => { });
            var factory = new TokenCredentialFactory(
                _ => { },
                (_, _, authenticationRecord) =>
                {
                    observed = authenticationRecord;
                    return new CountingCredential();
                },
                new AuthenticationRecordStore(path));

            factory.GetCredential(new GraphAuthOptions
            {
                AuthMode = AuthMode.Delegated,
                TenantId = "different-tenant-id",
                ClientId = record.ClientId,
                FallbackToMemoryTokenCache = false
            });

            observed.Should().BeNull();
        }
        finally
        {
            File.Delete(path);
            Directory.Delete(directory);
        }
    }

    [Theory]
    [InlineData("common")]
    [InlineData("consumers")]
    [InlineData("organizations")]
    public void New_factory_instance_accepts_special_tenant_aliases(string tenantId)
    {
        string directory = Path.Combine(Path.GetTempPath(), $"microsoft-mcp-{Guid.NewGuid():N}");
        string path = Path.Combine(directory, "authentication-record.json");
        AuthenticationRecord record = CreateAuthenticationRecord();
        AuthenticationRecord? observed = null;
        try
        {
            new AuthenticationRecordStore(path).Save(record, _ => { });
            var factory = new TokenCredentialFactory(
                _ => { },
                (_, _, authenticationRecord) =>
                {
                    observed = authenticationRecord;
                    return new CountingCredential();
                },
                new AuthenticationRecordStore(path));

            factory.GetCredential(new GraphAuthOptions
            {
                AuthMode = AuthMode.Delegated,
                TenantId = tenantId,
                ClientId = record.ClientId,
                FallbackToMemoryTokenCache = false
            });

            observed.Should().NotBeNull();
        }
        finally
        {
            File.Delete(path);
            Directory.Delete(directory);
        }
    }

    [Fact]
    public void Authentication_record_store_reads_through_file_system_abstraction()
    {
        const string path = "/tmp/authentication-record.json";
        var fileSystem = Substitute.For<IFileSystem>();
        var file = Substitute.For<IFile>();
        fileSystem.File.Returns(file);
        file.Exists(path).Returns(false);
        var store = new AuthenticationRecordStore(path, fileSystem);

        store.Load().Should().BeNull();
        file.Received().Exists(path);
    }

    [Fact]
    public void Authentication_record_store_injects_unix_permission_changes()
    {
        string directory = Path.Combine(Path.GetTempPath(), $"microsoft-mcp-{Guid.NewGuid():N}");
        string path = Path.Combine(directory, "authentication-record.json");
        List<(string Path, UnixFileMode Mode)> permissionChanges = [];
        try
        {
            var store = new AuthenticationRecordStore(
                path,
                setUnixFileMode: (permissionPath, mode) => permissionChanges.Add((permissionPath, mode)));

            store.Save(CreateAuthenticationRecord(), _ => { });

            permissionChanges.Count.Should().Be(OperatingSystem.IsWindows() ? 0 : 3);
        }
        finally
        {
            File.Delete(path);
            Directory.Delete(directory);
        }
    }

    private static AuthenticationRecord CreateAuthenticationRecord() =>
        IdentityModelFactory.AuthenticationRecord(
            "user@example.test",
            "https://login.microsoftonline.com/",
            "home-account-id",
            "tenant-id",
            "client-id");

    [Fact]
    public void Unencrypted_cache_warning_is_written_once()
    {
        List<string> warnings = [];
        var factory = new TokenCredentialFactory(warnings.Add);
        var options = new GraphAuthOptions
        {
            AuthMode = AuthMode.Delegated,
            DelegatedFlow = DelegatedFlow.DeviceCode,
            TenantId = "t",
            ClientId = "c",
            UnsafeAllowUnencryptedTokenCache = true
        };

        factory.GetCredential(options);
        factory.GetCredential(options);

        warnings.Count(w => w.Contains("unencrypted token-cache", StringComparison.OrdinalIgnoreCase)).Should().Be(1);
    }

    private sealed class FailingCredential(string message) : TokenCredential
    {
        public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken) =>
            throw new InvalidOperationException(message);

        public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken) =>
            ValueTask.FromException<AccessToken>(new InvalidOperationException(message));
    }

    private sealed class TimeoutCredential : TokenCredential
    {
        public override AccessToken GetToken(
            TokenRequestContext requestContext,
            CancellationToken cancellationToken) =>
            throw new TimeoutException("unrelated timeout");

        public override ValueTask<AccessToken> GetTokenAsync(
            TokenRequestContext requestContext,
            CancellationToken cancellationToken) =>
            ValueTask.FromException<AccessToken>(new TimeoutException("unrelated timeout"));
    }

    private sealed class HangingCredential : TokenCredential
    {
        private readonly TaskCompletionSource<AccessToken> _completion =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _calls;

        public int Calls => Volatile.Read(ref _calls);

        public void Release() => _completion.TrySetResult(
            new AccessToken("released", DateTimeOffset.UtcNow.AddMinutes(5)));

        public override AccessToken GetToken(
            TokenRequestContext requestContext,
            CancellationToken cancellationToken)
        {
            Interlocked.Increment(ref _calls);
            return _completion.Task.GetAwaiter().GetResult();
        }

        public override ValueTask<AccessToken> GetTokenAsync(
            TokenRequestContext requestContext,
            CancellationToken cancellationToken)
        {
            Interlocked.Increment(ref _calls);
            return new(_completion.Task);
        }
    }

    private sealed class CountingCredential : TokenCredential
    {
        public int Calls { get; private set; }

        public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            Calls++;
            return new AccessToken("token", DateTimeOffset.UtcNow.AddMinutes(5));
        }

        public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            Calls++;
            return ValueTask.FromResult(new AccessToken("token", DateTimeOffset.UtcNow.AddMinutes(5)));
        }
    }

    private sealed class AuthenticationRequiredCredential : TokenCredential
    {
        public bool Authenticated { get; set; }
        public string? FailureAfterAuthentication { get; set; }

        public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            if (!Authenticated)
            {
                throw new AuthenticationRequiredException("Authentication required.", requestContext);
            }

            if (FailureAfterAuthentication is not null)
            {
                throw new InvalidOperationException(FailureAfterAuthentication);
            }

            return new AccessToken("token", DateTimeOffset.UtcNow.AddMinutes(5));
        }

        public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            if (!Authenticated)
            {
                return ValueTask.FromException<AccessToken>(new AuthenticationRequiredException(
                    "Authentication required.",
                    requestContext));
            }

            return FailureAfterAuthentication is null
                ? ValueTask.FromResult(new AccessToken("token", DateTimeOffset.UtcNow.AddMinutes(5)))
                : ValueTask.FromException<AccessToken>(new InvalidOperationException(FailureAfterAuthentication));
        }
    }
}
