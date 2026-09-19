using Microsoft.Extensions.Options;

namespace MicrosoftMcp.Common;

public sealed class CalendarPolicyValidator : IValidateOptions<CalendarPolicyOptions>
{
    public ValidateOptionsResult Validate(string? name, CalendarPolicyOptions options)
    {
        if (options.RequireInternalAttendees
            && options.AllowedAttendeeDomains.Length == 0
            && options.AllowedAttendeeAddresses.Length == 0)
        {
            return ValidateOptionsResult.Fail(
                "Calendar policy: RequireInternalAttendees is true but no allowed attendee domains or exact addresses are configured. " +
                "Next: list internal attendee domains or exact addresses in the admin-owned policy.json.");
        }

        return ValidateOptionsResult.Success;
    }
}

public static class PolicySetValidator
{
    public static string? Validate(EffectivePolicySet policies)
    {
        ValidateOptionsResult outlook = new MessagingPolicyValidator().Validate(null, policies.Outlook);
        if (outlook.Failed)
        {
            return outlook.FailureMessage;
        }

        ValidateOptionsResult calendar = new CalendarPolicyValidator().Validate(null, policies.Calendar);
        return calendar.Failed ? calendar.FailureMessage : null;
    }
}
