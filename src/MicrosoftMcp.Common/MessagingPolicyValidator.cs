using Microsoft.Extensions.Options;

namespace MicrosoftMcp.Common;

public sealed class MessagingPolicyValidator : IValidateOptions<MessagingPolicyOptions>
{
    public ValidateOptionsResult Validate(string? name, MessagingPolicyOptions options)
    {
        if (options.RequireInternalRecipients && options.AllowedRecipientDomains.Length == 0)
        {
            return ValidateOptionsResult.Fail(
                "Messaging policy: RequireInternalRecipients is true but AllowedRecipientDomains is empty. " +
                "Next: list the internal domains in the admin-owned policy.json.");
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
