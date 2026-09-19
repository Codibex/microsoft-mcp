using System.Text.Json.Serialization;

namespace MicrosoftMcp.Common;

/// <summary>Shared address restrictions used by messaging services.</summary>
public class RecipientPolicyOptions
{
    public bool RequireInternalRecipients { get; set; }

    public string[] AllowedRecipientDomains { get; set; } = [];

    public string[] AllowedRecipientAddresses { get; set; } = [];

    [JsonIgnore]
    public virtual bool IsRestrictive => RequireInternalRecipients;
}

/// <summary>Shared recipient restrictions and AI disclosure settings.</summary>
public class DisclosurePolicyOptions : RecipientPolicyOptions
{
    public bool AiDisclosureEnabled { get; set; }

    public string AiDisclosureText { get; set; } = string.Empty;

    [JsonIgnore]
    public override bool IsRestrictive => base.IsRestrictive || AiDisclosureEnabled;
}

/// <summary>Outlook policy: recipient restrictions plus mail disclosure.</summary>
public class OutlookPolicyOptions : DisclosurePolicyOptions
{
}

/// <summary>Calendar policy with attendee-specific names and no mail disclosure settings.</summary>
public sealed class CalendarPolicyOptions
{
    public bool RequireInternalAttendees { get; set; }

    public string[] AllowedAttendeeDomains { get; set; } = [];

    public string[] AllowedAttendeeAddresses { get; set; } = [];

    [JsonIgnore]
    public bool IsRestrictive => RequireInternalAttendees;
}

/// <summary>Teams policy: conversation-member restrictions plus disclosure.</summary>
public sealed class TeamsPolicyOptions : DisclosurePolicyOptions
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

    public bool IsRestrictive => Outlook.IsRestrictive || Calendar.IsRestrictive || Teams.IsRestrictive;
}
