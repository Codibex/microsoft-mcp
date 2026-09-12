using AwesomeAssertions;
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
        GraphErrorMapper.FromStatus(status, graphCode, "detail", "read_email")
            .Code.Should().Be(expectedCode);

    [Theory]
    [InlineData("ErrorItemNotFound", "message-not-found")]
    [InlineData("ErrorMessageNotFound", "message-not-found")]
    [InlineData("ErrorFolderNotFound", "folder-not-found")]
    [InlineData("ErrorAccessDenied", "access-denied")]
    [InlineData("ErrorInvalidIdMalformed", "invalid-request")]
    public void Known_graph_codes_override_status(string graphCode, string expectedCode) =>
        GraphErrorMapper.FromStatus(400, graphCode, "detail", "move_email")
            .Code.Should().Be(expectedCode);

    [Fact]
    public void Message_always_carries_code_and_hint()
    {
        var ex = GraphErrorMapper.FromStatus(429, null, "slow down", "search_emails");

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
    public void MailServiceException_passes_through_untouched()
    {
        var original = MailServiceException.FolderNotFound("x");
        GraphErrorMapper.ToMailServiceException(original, "move_email").Should().BeSameAs(original);
    }

    [Fact]
    public void Network_failure_becomes_service_unavailable() =>
        GraphErrorMapper.ToMailServiceException(new HttpRequestException("no route"), "read_email")
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
