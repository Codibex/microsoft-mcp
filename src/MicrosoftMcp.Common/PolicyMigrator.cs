using System.Text;
using Microsoft.Extensions.Options;

namespace MicrosoftMcp.Common;

public sealed record PolicyMigrationResult(
    string PolicyPath,
    bool MigrationRequired,
    bool Written,
    string? BackupPath,
    int TargetVersion);

/// <summary>Explicit, atomic persistence of the legacy policy migration.</summary>
public static class PolicyMigrator
{
    public static PolicyMigrationResult Migrate(string path, bool write)
    {
        string normalizedPath = Path.GetFullPath(path);
        EffectivePolicySet policies = PolicyDocument.Load(normalizedPath);
        if (!policies.MigrationRequired)
        {
            return new PolicyMigrationResult(normalizedPath, false, false, null, 1);
        }

        string? validationError = PolicySetValidator.Validate(policies);
        if (validationError is not null)
        {
            throw new OptionsValidationException(
                "Messaging", typeof(EffectivePolicySet), [validationError]);
        }

        string content = PolicyDocument.SerializeVersioned(policies);
        if (!write)
        {
            return new PolicyMigrationResult(normalizedPath, true, false, null, 1);
        }

        PolicyFile.EnsureMigrationWritable(normalizedPath);
        PolicyFile.EnsureProtected(normalizedPath, policies);

        string directory = Path.GetDirectoryName(normalizedPath)
            ?? throw new IOException($"Policy path '{path}' has no parent directory.");
        string backup = normalizedPath + $".legacy.{DateTime.UtcNow:yyyyMMddHHmmss}.{Guid.NewGuid():N}.bak";
        string temporary = Path.Combine(
            directory,
            $".{Path.GetFileName(normalizedPath)}.migration-{Guid.NewGuid():N}.tmp");
        bool replaced = false;

        try
        {
            File.Copy(normalizedPath, backup, overwrite: false);
            WriteAndFlush(temporary, content);
            PreserveUnixMode(normalizedPath, temporary);

            if (OperatingSystem.IsWindows())
            {
                File.Replace(temporary, normalizedPath, null, ignoreMetadataErrors: true);
            }
            else
            {
                File.Move(temporary, normalizedPath, overwrite: true);
            }

            replaced = true;
            PolicyFile.EnsureProtected(normalizedPath, policies);
            return new PolicyMigrationResult(normalizedPath, true, true, backup, 1);
        }
        catch
        {
            if (replaced && File.Exists(backup))
            {
                File.Move(backup, normalizedPath, overwrite: true);
            }

            throw;
        }
        finally
        {
            if (File.Exists(temporary))
            {
                File.Delete(temporary);
            }
        }
    }

    private static void WriteAndFlush(string path, string content)
    {
        byte[] bytes = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(content);
        using FileStream stream = new(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        stream.Write(bytes);
        stream.Flush(flushToDisk: true);
    }

    private static void PreserveUnixMode(string source, string destination)
    {
        if (!OperatingSystem.IsWindows())
        {
            File.SetUnixFileMode(destination, File.GetUnixFileMode(source));
        }
    }
}
