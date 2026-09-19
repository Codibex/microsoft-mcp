using System.Text.Json;
using MicrosoftMcp.Common;
using MicrosoftMcp.Host.Setup;

namespace MicrosoftMcp.Host.Tests;

public sealed class PolicyMigrationCliTests
{
    private static readonly object ConsoleLock = new();

    [Fact]
    public void Policy_migrate_json_supports_preview_and_write()
    {
        string path = Path.Combine(Path.GetTempPath(), $"policy-cli-{Guid.NewGuid():N}.json");
        string? backupPath = null;
        try
        {
            File.WriteAllText(path, """{"requireInternalRecipients":false}""");

            (int previewCode, string previewOutput) = Run(
                "policy", "migrate", "--path", path, "--json");
            Assert.Equal(0, previewCode);
            using (JsonDocument preview = JsonDocument.Parse(previewOutput))
            {
                Assert.True(preview.RootElement.GetProperty("migrationRequired").GetBoolean());
                Assert.False(preview.RootElement.GetProperty("written").GetBoolean());
            }

            (int writeCode, string writeOutput) = Run(
                "policy", "migrate", "--path", path, "--write", "--json");
            Assert.Equal(0, writeCode);
            using (JsonDocument written = JsonDocument.Parse(writeOutput))
            {
                Assert.True(written.RootElement.GetProperty("written").GetBoolean());
                backupPath = written.RootElement.GetProperty("backupPath").GetString();
            }

            Assert.NotNull(backupPath);
            Assert.True(File.Exists(backupPath));
            Assert.False(PolicyDocument.Load(path).MigrationRequired);
        }
        finally
        {
            File.Delete(path);
            if (backupPath is not null)
            {
                File.Delete(backupPath);
            }
        }
    }

    [Fact]
    public void Policy_migrate_rejects_unknown_action()
    {
        (int code, string output) = Run("policy", "unknown", "--json");

        Assert.Equal(2, code);
        using JsonDocument result = JsonDocument.Parse(output);
        Assert.Contains("policy migrate", result.RootElement.GetProperty("error").GetString());
    }

    private static (int Code, string Output) Run(params string[] args)
    {
        lock (ConsoleLock)
        {
            TextWriter original = Console.Out;
            using StringWriter output = new();
            try
            {
                Console.SetOut(output);
                return (SetupCli.Run(args), output.ToString());
            }
            finally
            {
                Console.SetOut(original);
            }
        }
    }
}
