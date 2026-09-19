namespace MicrosoftMcp.Common;

/// <summary>Shared address restrictions used by messaging services.</summary>
public class RecipientPolicyOptions
{
    public bool RequireInternalRecipients { get; set; }

    public string[] AllowedRecipientDomains { get; set; } = [];

    public string[] AllowedRecipientAddresses { get; set; } = [];

    public virtual bool IsRestrictive => RequireInternalRecipients;
}

/// <summary>Outlook policy: recipient restrictions plus mail disclosure.</summary>
public class OutlookPolicyOptions : RecipientPolicyOptions
{
    public bool AiDisclosureEnabled { get; set; }

    public string AiDisclosureText { get; set; } = string.Empty;

    public override bool IsRestrictive => base.IsRestrictive || AiDisclosureEnabled;
}

/// <summary>Calendar policy with attendee-specific names and no mail disclosure settings.</summary>
public sealed class CalendarPolicyOptions
{
    public bool RequireInternalAttendees { get; set; }

    public string[] AllowedAttendeeDomains { get; set; } = [];

    public string[] AllowedAttendeeAddresses { get; set; } = [];

    public bool IsRestrictive => RequireInternalAttendees;
}

/// <summary>Reserved for future Teams write-policy rules.</summary>
public sealed class TeamsPolicyOptions
{
}

/// <summary>Effective policies derived from one admin-owned policy document.</summary>
public sealed record EffectivePolicySet(
    OutlookPolicyOptions Outlook,
    CalendarPolicyOptions Calendar,
    TeamsPolicyOptions Teams,
    int? SourceVersion,
    bool IsLegacy)
{
    public bool MigrationRequired => IsLegacy;

    public bool IsRestrictive => Outlook.IsRestrictive || Calendar.IsRestrictive;
}
