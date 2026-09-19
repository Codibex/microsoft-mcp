using System.Text.Json;
using Microsoft.Extensions.Options;

namespace MicrosoftMcp.Common;

/// <summary>One-line startup wiring shared by all hosts: find, load, validate,
/// protection-check and audit-log the admin-owned policy.json. The policy file is
/// the ONLY source for these settings — Messaging__* env vars are ignored by design
/// (mcp.json is user-writable, so env must not override admin policy).</summary>
public static class MessagingPolicySetup
{
    public static EffectivePolicySet InitializePolicies()
    {
        string? path = PolicyFile.FindPolicyFile();
        if (path is null)
        {
            Console.Error.WriteLine(
                "[startup] No policy.json found (system path or next to the binary) – " +
                "recipient/disclosure policy disabled.");
            return new EffectivePolicySet(
                new OutlookPolicyOptions(),
                new CalendarPolicyOptions(),
                new TeamsPolicyOptions(),
                null,
                false);
        }

        EffectivePolicySet policies;
        try
        {
            policies = PolicyDocument.Load(path);
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
        {
            throw new OptionsValidationException(
                "Messaging", typeof(EffectivePolicySet), [$"Policy file '{path}' could not be read: {ex.Message}"]);
        }

        string? validationError = PolicySetValidator.Validate(policies);
        if (validationError is not null)
        {
            throw new OptionsValidationException(
                "Messaging", typeof(EffectivePolicySet), [validationError]);
        }

        PolicyFile.EnsureProtected(path, policies);
        Console.Error.WriteLine(
            $"[startup] Policy: {path} (sha256 {PolicyFile.ComputeHash(path)[..12]}…) | " +
            $"format={(policies.IsLegacy ? "legacy" : "v1")} " +
            $"outlook-internal: {policies.Outlook.RequireInternalRecipients} " +
            $"calendar-internal: {policies.Calendar.RequireInternalAttendees} " +
            $"disclosure: {policies.Outlook.AiDisclosureEnabled}");
        return policies;
    }

    /// <summary>Compatibility entry point returning the effective Outlook policy.</summary>
    public static MessagingPolicyOptions Initialize()
    {
        OutlookPolicyOptions outlook = InitializePolicies().Outlook;
        return new MessagingPolicyOptions
        {
            RequireInternalRecipients = outlook.RequireInternalRecipients,
            AllowedRecipientDomains = [.. outlook.AllowedRecipientDomains],
            AllowedRecipientAddresses = [.. outlook.AllowedRecipientAddresses],
            AiDisclosureEnabled = outlook.AiDisclosureEnabled,
            AiDisclosureText = outlook.AiDisclosureText
        };
    }
}
