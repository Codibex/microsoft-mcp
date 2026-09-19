using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using MicrosoftMcp.Calendar;
using MicrosoftMcp.Common;
using MicrosoftMcp.Outlook;
using MicrosoftMcp.Teams;

namespace MicrosoftMcp.Host.Tests;

public sealed class PolicyInjectionTests
{
    [Fact]
    public void Service_registrations_keep_policy_sections_typed_and_separate()
    {
        var outlook = new OutlookPolicyOptions
        {
            RequireInternalRecipients = true,
            AllowedRecipientDomains = ["mail.example"]
        };
        var calendar = new CalendarPolicyOptions
        {
            RequireInternalAttendees = true,
            AllowedAttendeeDomains = ["calendar.example"]
        };
        var teams = new TeamsPolicyOptions();

        using ServiceProvider provider = new ServiceCollection()
            .AddOutlook(outlook)
            .AddCalendar(calendar)
            .AddTeams(teams)
            .BuildServiceProvider();

        Assert.Same(outlook, provider.GetRequiredService<IOptions<OutlookPolicyOptions>>().Value);
        Assert.Same(calendar, provider.GetRequiredService<IOptions<CalendarPolicyOptions>>().Value);
        Assert.Same(teams, provider.GetRequiredService<IOptions<TeamsPolicyOptions>>().Value);
    }

    [Fact]
    public void Calendar_registration_accepts_the_legacy_shared_policy()
    {
        var legacy = new MessagingPolicyOptions
        {
            RequireInternalRecipients = true,
            AllowedRecipientDomains = ["legacy.example"],
            AllowedRecipientAddresses = ["guest@example.com"]
        };

        using ServiceProvider provider = new ServiceCollection()
            .AddCalendar(legacy)
            .BuildServiceProvider();

        CalendarPolicyOptions policy = provider
            .GetRequiredService<IOptions<CalendarPolicyOptions>>()
            .Value;
        Assert.True(policy.RequireInternalAttendees);
        Assert.Equal(["legacy.example"], policy.AllowedAttendeeDomains);
        Assert.Equal(["guest@example.com"], policy.AllowedAttendeeAddresses);
    }
}
