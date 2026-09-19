using System.Net;
using System.Text;
using System.Text.Json;
using AwesomeAssertions;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Kiota.Http.HttpClientLibrary;
using MicrosoftMcp.Common;
using NSubstitute;

namespace MicrosoftMcp.Calendar.Tests;

public sealed class GraphCalendarServiceTests
{
    [Fact]
    public async Task CreateEvent_posts_to_named_calendar_with_write_fields()
    {
        var handler = new RecordingHandler(
            "{\"id\":\"event-1\",\"subject\":\"Planning\",\"start\":{\"dateTime\":\"2026-09-14T12:00:00\",\"timeZone\":\"UTC\"},\"end\":{\"dateTime\":\"2026-09-14T13:00:00\",\"timeZone\":\"UTC\"}}");
        var service = CreateService(handler);

        var created = await service.CreateEventAsync(
            "Planning",
            "2026-09-14T14:00:00+02:00",
            "2026-09-14T15:00:00+02:00",
            "calendar-1",
            "Agenda",
            "Room 4",
            ["person@example.com"],
            reminderMinutesBeforeStart: 15);

        created.Id.Should().Be("event-1");
        handler.Method.Should().Be(HttpMethod.Post);
        handler.RequestUri!.AbsolutePath.Should().Be("/v1.0/me/calendars/calendar-1/events");

        using var json = JsonDocument.Parse(handler.Body!);
        json.RootElement.GetProperty("subject").GetString().Should().Be("Planning");
        json.RootElement.GetProperty("body").GetProperty("content").GetString().Should().Be("Agenda");
        json.RootElement.GetProperty("location").GetProperty("displayName").GetString().Should().Be("Room 4");
        json.RootElement.GetProperty("attendees")[0].GetProperty("emailAddress")
            .GetProperty("address").GetString().Should().Be("person@example.com");
        json.RootElement.GetProperty("reminderMinutesBeforeStart").GetInt32().Should().Be(15);
    }

    [Fact]
    public async Task UpdateEvent_patches_default_event_with_only_supplied_fields()
    {
        var handler = new RecordingHandler("{\"id\":\"event-1\",\"subject\":\"Moved\"}");
        var service = CreateService(handler);

        var updated = await service.UpdateEventAsync("event-1", subject: "Moved");

        updated.Id.Should().Be("event-1");
        handler.Method.Should().Be(HttpMethod.Patch);
        handler.RequestUri!.AbsolutePath.Should().Be("/v1.0/me/calendar/events/event-1");

        using var json = JsonDocument.Parse(handler.Body!);
        json.RootElement.GetProperty("subject").GetString().Should().Be("Moved");
        json.RootElement.TryGetProperty("start", out _).Should().BeFalse();
        json.RootElement.TryGetProperty("end", out _).Should().BeFalse();
    }

    [Fact]
    public async Task UpdateEvent_revalidates_existing_attendees_when_not_replaced()
    {
        var handler = new RecordingHandler(
            "{\"id\":\"event-1\",\"attendees\":[{\"emailAddress\":{\"address\":\"outside@example.com\"}}]}");
        var policy = new CalendarPolicyOptions
        {
            RequireInternalAttendees = true,
            AllowedAttendeeDomains = ["firma.de"]
        };
        var service = CreateService(handler, policy: policy);

        var act = () => service.UpdateEventAsync("event-1", subject: "Moved");

        await act.Should().ThrowAsync<GraphServiceException>()
            .WithMessage("*[invalid-request]*policy.json*");
        handler.Method.Should().Be(HttpMethod.Get);
    }

    [Fact]
    public async Task UpdateEvent_rejects_partial_time_update_that_would_invert_window()
    {
        var handler = new RecordingHandler(
            "{\"id\":\"event-1\",\"start\":{\"dateTime\":\"2026-09-14T14:00:00\",\"timeZone\":\"UTC\"},\"end\":{\"dateTime\":\"2026-09-14T15:00:00\",\"timeZone\":\"UTC\"}}");
        var service = CreateService(handler);

        var act = () => service.UpdateEventAsync(
            "event-1", start: "2026-09-14T16:00:00Z");

        await act.Should().ThrowAsync<GraphServiceException>()
            .WithMessage("*[invalid-request]*end*after*start*");
        handler.Method.Should().Be(HttpMethod.Get);
    }

    [Fact]
    public async Task CreateEvent_preserves_local_date_for_all_day_values_with_offset()
    {
        var handler = new RecordingHandler("{\"id\":\"event-1\"}");
        var service = CreateService(handler);

        await service.CreateEventAsync(
            "Holiday",
            "2026-09-14T00:00:00+02:00",
            "2026-09-15T00:00:00+02:00",
            isAllDay: true);

        using var json = JsonDocument.Parse(handler.Body!);
        json.RootElement.GetProperty("start").GetProperty("dateTime").GetString()
            .Should().Be("2026-09-14T00:00:00");
        json.RootElement.GetProperty("end").GetProperty("dateTime").GetString()
            .Should().Be("2026-09-15T00:00:00");
        json.RootElement.GetProperty("start").GetProperty("timeZone").GetString()
            .Should().Be("UTC");
    }

    [Fact]
    public async Task CreateEvent_rejects_non_midnight_all_day_values()
    {
        var handler = new RecordingHandler("{}");
        var service = CreateService(handler);

        var act = () => service.CreateEventAsync(
            "Holiday",
            "2026-09-14T01:00:00+02:00",
            "2026-09-15T01:00:00+02:00",
            isAllDay: true);

        await act.Should().ThrowAsync<GraphServiceException>()
            .WithMessage("*[invalid-request]*midnight*");
        handler.Method.Should().BeNull();
    }

    [Fact]
    public async Task Calendar_writes_reject_app_only_auth()
    {
        var handler = new RecordingHandler("{}");
        var options = new GraphAuthOptions { AuthMode = AuthMode.AppOnly, UserIdOrUpn = "user@example.com" };
        var service = CreateService(handler, options);

        var act = () => service.CreateEventAsync(
            "Planning", "2026-09-14T14:00:00Z", "2026-09-14T15:00:00Z");

        await act.Should().ThrowAsync<GraphServiceException>()
            .WithMessage("*[auth-misconfigured]*");
        handler.Method.Should().BeNull();
    }

    [Fact]
    public async Task CreateEvent_rejects_external_attendee_when_policy_requires_internal()
    {
        var handler = new RecordingHandler("{}");
        var policy = new CalendarPolicyOptions
        {
            RequireInternalAttendees = true,
            AllowedAttendeeDomains = ["firma.de"]
        };
        var service = CreateService(handler, policy: policy);

        var act = () => service.CreateEventAsync(
            "Planning",
            "2026-09-14T14:00:00Z",
            "2026-09-14T15:00:00Z",
            attendees: ["outside@example.com"]);

        await act.Should().ThrowAsync<GraphServiceException>()
            .WithMessage("*[invalid-request]*policy.json*");
        handler.Method.Should().BeNull();
    }

    [Fact]
    public async Task UpdateEvent_rejects_external_attendee_before_patch()
    {
        var handler = new RecordingHandler("{}");
        var policy = new CalendarPolicyOptions
        {
            RequireInternalAttendees = true,
            AllowedAttendeeDomains = ["firma.de"]
        };
        var service = CreateService(handler, policy: policy);

        var act = () => service.UpdateEventAsync(
            "event-1", attendees: ["outside@example.com"]);

        await act.Should().ThrowAsync<GraphServiceException>()
            .WithMessage("*[invalid-request]*policy.json*");
        handler.Method.Should().BeNull();
    }

    private static GraphCalendarService CreateService(
        RecordingHandler handler,
        GraphAuthOptions? options = null,
        CalendarPolicyOptions? policy = null)
    {
        var httpClient = new HttpClient(handler);
        var requestAdapter = new HttpClientRequestAdapter(
            Substitute.For<Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider>(),
            httpClient: httpClient);
        var graphClient = new GraphServiceClient(requestAdapter);
        return new GraphCalendarService(
            graphClient,
            Options.Create(options ?? new GraphAuthOptions()),
            Options.Create(policy ?? new CalendarPolicyOptions()));
    }

    private sealed class RecordingHandler(string response) : HttpMessageHandler
    {
        public HttpMethod? Method { get; private set; }
        public Uri? RequestUri { get; private set; }
        public string? Body { get; private set; }

        protected override async Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request, CancellationToken cancellationToken)
        {
            Method = request.Method;
            RequestUri = request.RequestUri;
            Body = request.Content is null
                ? null
                : await request.Content.ReadAsStringAsync(cancellationToken);
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(response, Encoding.UTF8, "application/json")
            };
        }
    }
}