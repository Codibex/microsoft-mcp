using Microsoft.Extensions.Options;

namespace MicrosoftMcp.Common;

public sealed class MessagingPolicyValidator : IValidateOptions<DisclosurePolicyOptions>
{
    public ValidateOptionsResult Validate(string? name, DisclosurePolicyOptions options)
    {
        if (options.RequireInternalRecipients
            && options.AllowedRecipientDomains.Length == 0
            && options.AllowedRecipientAddresses.Length == 0)
        {
            return ValidateOptionsResult.Fail(
                "Messaging policy: RequireInternalRecipients is true but no allowed domains or exact addresses are configured. " +
                "Next: list internal domains or exact addresses in the admin-owned policy.json.");
        }

        if (options.AiDisclosureEnabled && string.IsNullOrWhiteSpace(options.AiDisclosureText))
        {
            return ValidateOptionsResult.Fail(
                "Messaging policy: AiDisclosureEnabled is true but AiDisclosureText is empty. " +
                "Next: set the disclosure text in the admin-owned policy.json.");
        }

        return ValidateOptionsResult.Success;
    }
}
