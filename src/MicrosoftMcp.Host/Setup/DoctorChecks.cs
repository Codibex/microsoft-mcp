using MicrosoftMcp.Common;

namespace MicrosoftMcp.Host.Setup;

/// <summary>One offline check with a fix hint.</summary>
public sealed record SetupCheck(string Id, bool Ok, string Message, string? Next);

/// <summary>Offline checks over the host config (no network, no token).</summary>
public static class DoctorChecks
{
    private static readonly string[] SpecialTenants = ["common", "consumers", "organizations"];

    /// <summary>Runs all checks. Policy: missing = ok (disabled), unreadable = fail.</summary>
    public static IReadOnlyList<SetupCheck> Run(
        IReadOnlyList<string> servers,
        GraphAuthOptions options,
        string? policyPath,
        string? policyError)
    {
        List<SetupCheck> checks = [];

        if (!string.IsNullOrWhiteSpace(policyError))
        {
            checks.Add(new SetupCheck("policy", false, $"Policy: {policyError}", "Next: grant read access or remove the unreadable file (fail-closed, see docs/outlook.md §10)."));
        }
        else if (policyPath is null)
        {
            bool needsPolicy = servers.Contains("outlook", StringComparer.OrdinalIgnoreCase)
                || servers.Contains("teams", StringComparer.OrdinalIgnoreCase);
            checks.Add(new SetupCheck(
                "policy",
                true,
                needsPolicy
                    ? "No policy.json found – recipient/disclosure policy disabled."
                    : "No policy.json found – not needed for these servers.",
                needsPolicy ? "Next (optional): deploy admin-owned policy.json, see docs/outlook.md §10." : null));
        }
        else
        {
            checks.Add(new SetupCheck("policy", true, $"Policy found: {policyPath}", null));
        }

        if (string.IsNullOrWhiteSpace(options.TenantId)
            || options.TenantId.Contains("YOUR-", StringComparison.OrdinalIgnoreCase))
        {
            checks.Add(new SetupCheck("tenant", false, "Graph:TenantId is missing or still a placeholder.", "Next: set Graph__TenantId (work: tenant GUID, personal: common)."));
        }
        else if (SpecialTenants.Contains(options.TenantId, StringComparer.OrdinalIgnoreCase)
            || Guid.TryParse(options.TenantId, out _))
        {
            checks.Add(new SetupCheck("tenant", true, $"TenantId looks plausible ({options.TenantId}).", null));
        }
        else
        {
            checks.Add(new SetupCheck("tenant", false, $"TenantId '{options.TenantId}' is neither a GUID nor common/consumers/organizations.", "Next: work → tenant GUID, personal → common."));
        }

        if (string.IsNullOrWhiteSpace(options.ClientId)
            || options.ClientId.Contains("YOUR-", StringComparison.OrdinalIgnoreCase))
        {
            checks.Add(new SetupCheck("client", false, "Graph:ClientId is missing or still a placeholder.", "Next: set Graph__ClientId from the Entra app overview."));
        }
        else if (Guid.TryParse(options.ClientId, out _))
        {
            checks.Add(new SetupCheck("client", true, "ClientId is a GUID.", null));
        }
        else
        {
            checks.Add(new SetupCheck("client", false, $"ClientId '{options.ClientId}' is not a GUID.", "Next: copy the Application (client) ID from Entra."));
        }

        string? comboError = SetupGuide.ValidateCombination(servers, options.AuthMode);
        checks.Add(comboError is null
            ? new SetupCheck("authmode", true, $"AuthMode {options.AuthMode} fits servers {string.Join(",", servers)}.", null)
            : new SetupCheck("authmode", false, $"AuthMode {options.AuthMode} does not fit servers {string.Join(",", servers)}.", comboError));

        checks.Add(CacheCheck(options));

        if (options.AuthMode == AuthMode.AppOnly)
        {
            if (string.Equals(options.UserIdOrUpn, "me", StringComparison.OrdinalIgnoreCase))
            {
                checks.Add(new SetupCheck("mailbox", false, "UserIdOrUpn is 'me' but AppOnly needs a mailbox.", "Next: set Graph__UserIdOrUpn to the mailbox UPN."));
            }
            else
            {
                checks.Add(new SetupCheck("mailbox", true, $"AppOnly mailbox: {options.UserIdOrUpn}.", null));
            }

            checks.Add(string.IsNullOrWhiteSpace(options.ClientSecret)
                ? new SetupCheck("secret", false, "ClientSecret is missing.", "Next: set Graph__ClientSecret via env (never commit).")
                : new SetupCheck("secret", true, "ClientSecret is set (value hidden).", null));
        }

        IReadOnlyList<string> required = SetupGuide.ScopesFor(servers);
        if (options.DelegatedScopes.Length == 0)
        {
            checks.Add(new SetupCheck("scopes", true, $"Scopes follow the selection: {string.Join(" ", required)}.", null));
        }
        else
        {
            string[] missing = [.. required.Where(s => !options.DelegatedScopes.Contains(s, StringComparer.OrdinalIgnoreCase))];
            checks.Add(missing.Length == 0
                ? new SetupCheck("scopes", true, $"Explicit scopes cover the selection: {string.Join(" ", options.DelegatedScopes)}.", null)
                : new SetupCheck("scopes", false, $"Explicit scopes miss: {string.Join(" ", missing)}.", "Next: add the missing delegated scopes in Entra + consent again."));
        }

        checks.Add(new SetupCheck(
            "flow",
            true,
            options.DelegatedFlow == DelegatedFlow.DeviceCode
                ? "DelegatedFlow DeviceCode (headless-ready)."
                : "DelegatedFlow Auto (browser login; headless needs DeviceCode).",
            options.DelegatedFlow == DelegatedFlow.DeviceCode
                ? null
                : "Next (headless only): set Graph__DelegatedFlow=DeviceCode."));

        return checks;
    }

    private static SetupCheck CacheCheck(GraphAuthOptions options)
    {
        if (options.AuthMode == AuthMode.AppOnly || !options.EnableTokenCache)
        {
            return new SetupCheck(
                "cache",
                true,
                options.EnableTokenCache
                    ? "Token cache is not used for AppOnly auth."
                    : "Persistent token cache disabled; a new login is required after each process start.",
                null);
        }

        if (options.UnsafeAllowUnencryptedTokenCache)
        {
            return new SetupCheck(
                "cache",
                true,
                "WARNING: token cache may be stored unencrypted on disk.",
                "Next: prefer the encrypted OS cache, or set Graph__EnableTokenCache=false.");
        }

        if (options.FallbackToMemoryTokenCache)
        {
            return new SetupCheck(
                "cache",
                true,
                "Encrypted OS token cache preferred; in-memory fallback enabled if Secret Service is unavailable.",
                "Next (optional): install GNOME Keyring/libsecret for persistence across restarts.");
        }

        bool likelyHeadlessLinux = OperatingSystem.IsLinux()
            && string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("DBUS_SESSION_BUS_ADDRESS"));
        return likelyHeadlessLinux
            ? new SetupCheck(
                "cache",
                false,
                "Persistent token cache requires a Linux Secret Service, but no D-Bus session was detected.",
                "Next: enable GNOME Keyring/libsecret, or set Graph__FallbackToMemoryTokenCache=true.")
            : new SetupCheck(
                "cache",
                true,
                "Encrypted OS token cache required; Secret Service availability is not verified offline.",
                "Next: set Graph__FallbackToMemoryTokenCache=true unless persistence is guaranteed.");
    }
}
