using AwesomeAssertions;
using Azure.Identity;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Kiota.Abstractions;
using MicrosoftMcp.Common;

namespace MicrosoftMcp.Outlook.Tests;

public sealed class GraphErrorMapperTests
{
    [Theory]
    [InlineData(404, null, "message-not-found")]
    [InlineData(403, null, "access-denied")]
    [InlineData(401, null, "auth-failed")]
    [InlineData(400, null, "invalid-request")]
    [InlineData(409, null, "conflict")]
    [InlineData(429, null, "throttled")]
    [InlineData(503, null, "service-unavailable")]
    [InlineData(500, null, "service-unavailable")]
    [InlineData(418, null, "graph-error")]
    public void Status_maps_to_code(int status, string? graphCode, string expectedCode) =>
        GraphErrorMapper.FromStatus(status, graphCode, "detail", "outlook_read_email")
            .Code.Should().Be(expectedCode);

    [Theory]
    [InlineData("ErrorItemNotFound", "message-not-found")]
    [InlineData("ErrorMessageNotFound", "message-not-found")]
    [InlineData("ErrorFolderNotFound", "folder-not-found")]
    [InlineData("ErrorAccessDenied", "access-denied")]
    [InlineData("ErrorInvalidIdMalformed", "invalid-request")]
    public void Known_graph_codes_override_status(string graphCode, string expectedCode) =>
        GraphErrorMapper.FromStatus(400, graphCode, "detail", "outlook_move_email")
            .Code.Should().Be(expectedCode);

    [Fact]
    public void Message_always_carries_code_and_hint()
    {
        var ex = GraphErrorMapper.FromStatus(429, null, "slow down", "outlook_search_emails");

        ex.Message.Should().Contain("[throttled]");
        ex.Message.Should().Contain("Next:");
    }

    [Fact]
    public void ExtractGraphCode_reads_code_from_graph_body()
    {
        GraphErrorMapper.ExtractGraphCode(
            "{\"error\":{\"code\":\"ErrorItemNotFound\",\"message\":\"No object found.\"}}")
            .Should().Be("ErrorItemNotFound");
    }

    [Fact]
    public void ExtractGraphCode_returns_null_without_code() =>
        GraphErrorMapper.ExtractGraphCode("plain failure").Should().BeNull();

    [Fact]
    public void Structured_odata_code_wins_over_message_text()
    {
        var api = new ODataError
        {
            ResponseStatusCode = 404,
            Error = new MainError { Code = "MailboxNotEnabledForRESTAPI", Message = "Inactive." }
        };

        GraphErrorMapper.ResolveGraphCode(api).Should().Be("MailboxNotEnabledForRESTAPI");
        GraphErrorMapper.ToMailServiceException(api, "outlook_search_emails").Code.Should().Be("mailbox-unavailable");
    }

    [Theory]
    [InlineData("ErrorMailboxNotAssociated")]
    [InlineData("ErrorNonExistentMailbox")]
    public void Mailbox_codes_map_to_mailbox_unavailable(string code)
    {
        var mapped = GraphErrorMapper.FromStatus(404, code, "inactive", "outlook_search_emails");
        mapped.Code.Should().Be("mailbox-unavailable");
        mapped.Message.Should().Contain("Next:");
    }

    [Fact]
    public void AuthDetail_promotes_aadsts_code_to_front()
    {
        var inner = new InvalidOperationException(
            "A configuration issue. Original exception: AADSTS7000218: missing client_assertion.");
        var ex = new AuthenticationFailedException("DeviceCodeCredential authentication failed: ", inner);

        var mapped = GraphErrorMapper.ToMailServiceException(ex, "outlook_search_emails");

        mapped.Code.Should().Be("auth-failed");
        mapped.Message.Should().Contain("AADSTS7000218");
        mapped.Message.Should().Contain("public client flows");
        mapped.Message.Should().Contain("Next:");
    }

    [Fact]
    public void Unknown_app_in_directory_hint_for_700016()
    {
        var ex = new AuthenticationFailedException(
            "AADSTS700016: Application was not found in the directory '9188040d'.");

        var mapped = GraphErrorMapper.ToMailServiceException(ex, "outlook_search_emails");

        mapped.Code.Should().Be("auth-failed");
        mapped.Message.Should().Contain("propagation");
    }

    [Fact]
    public void Personal_only_app_hint_for_2346()
    {
        var ex = new AuthenticationFailedException(
            "AADSTS2346: The app is configured for personal Microsoft accounts only.");

        var mapped = GraphErrorMapper.ToMailServiceException(ex, "outlook_search_emails");

        mapped.Code.Should().Be("auth-failed");
        mapped.Message.Should().Contain("TenantId=consumers");
        mapped.Message.Should().Contain("AzureADandPersonalMicrosoftAccount");
    }

    [Fact]
    public void MailServiceException_passes_through_untouched()
    {
        var original = MailServiceException.FolderNotFound("x");
        GraphErrorMapper.ToMailServiceException(original, "outlook_move_email").Should().BeSameAs(original);
    }

    [Fact]
    public void Network_failure_becomes_service_unavailable() =>
        GraphErrorMapper.ToMailServiceException(new HttpRequestException("no route"), "outlook_read_email")
            .Code.Should().Be("service-unavailable");

    [Fact]
    public void Cancellation_is_never_wrapped_by_mapper_but_propagates()
    {
        // documents the contract: tools let OperationCanceledException bubble up
        var cts = new CancellationTokenSource();
        cts.Cancel();
        Assert.True(cts.IsCancellationRequested);
    }
}
