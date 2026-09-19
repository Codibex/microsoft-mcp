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
}
