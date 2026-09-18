using System.Text.Json;
using System.Text.Json.Serialization;
using MicrosoftMcp.Common;
using ModelContextProtocol.Protocol;

namespace MicrosoftMcp.Teams;

/// <summary>Same contract as the other hosts: success as JSON text, failures
/// as isError results with "[code] ... Next: ...".</summary>
internal static class ToolResult
{
    internal static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    internal static CallToolResult Ok<T>(T payload) => new()
    {
        Content = [new TextContentBlock
        {
            Text = System.Text.Json.JsonSerializer.Serialize(payload, Json)
        }]
    };

    internal static CallToolResult Fail(GraphServiceException error) => new()
    {
        Content = [new TextContentBlock { Text = error.Message }],
        IsError = true
    };

    internal static string ReadText(CallToolResult result) =>
        result.Content.Count == 1 && result.Content[0] is TextContentBlock text
            ? text.Text
            : throw new InvalidOperationException("Unexpected tool result shape.");

    internal static T ReadJson<T>(CallToolResult result) =>
        System.Text.Json.JsonSerializer.Deserialize<T>(ReadText(result), Json)
            ?? throw new InvalidOperationException("Tool result was null JSON.");
}
