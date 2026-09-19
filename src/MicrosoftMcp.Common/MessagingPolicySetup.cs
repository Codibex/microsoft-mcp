using System.Text.Json;
using Microsoft.Extensions.Options;

namespace MicrosoftMcp.Common;

/// <summary>One-line startup wiring shared by all hosts: find, load, validate,
/// protection-check and audit-log the admin-owned policy.json. The policy file is
/// the ONLY source for these settings — Messaging__* env vars are ignored by design
/// (mcp.json is user-writable, so env must not override admin policy).</summary>
public static class MessagingPolicySetup
{
    public static MessagingPolicyOptions Initialize()
    {
        string? path = PolicyFile.FindPolicyFile();
        if (path is null)
        {
            Console.Error.WriteLine(
                "[startup] No policy.json found (system path or next to the binary) – " +
                "recipient/disclosure policy disabled.");
            return new MessagingPolicyOptions();
        }

        MessagingPolicyOptions policy;
        try
        {
            policy = PolicyFile.Load(path);
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
        {
            throw new OptionsValidationException(
                "Messaging", typeof(MessagingPolicyOptions), [$"Policy file '{path}' could not be read: {ex.Message}"]);
        }

        var result = new MessagingPolicyValidator().Validate(null, policy);
        if (result.Failed)
        {
            throw new OptionsValidationException(
                "Messaging", typeof(MessagingPolicyOptions), [result.FailureMessage ?? "Invalid policy."]);
        }

        PolicyFile.EnsureProtected(path, policy);
        Console.Error.WriteLine(
            $"[startup] Policy: {path} (sha256 {PolicyFile.ComputeHash(path)[..12]}…) | " +
            $"internal-only: {policy.RequireInternalRecipients} " +
            $"domains=[{string.Join(",", policy.AllowedRecipientDomains)}] " +
            $"addresses=[{string.Join(",", policy.AllowedRecipientAddresses)}] | " +
            $"disclosure: {policy.AiDisclosureEnabled}");
        return policy;
    }
}
