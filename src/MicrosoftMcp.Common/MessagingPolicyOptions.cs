namespace MicrosoftMcp.Common;

/// <summary>Compatibility type for callers using the pre-versioned policy model.
/// New code should use <see cref="OutlookPolicyOptions"/> and
/// <see cref="CalendarPolicyOptions"/> from the effective policy set.</summary>
public sealed class MessagingPolicyOptions : OutlookPolicyOptions
{
}
