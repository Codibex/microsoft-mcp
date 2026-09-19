using System.Text;

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
        EffectivePolicySet policies = PolicyDocument.Load(path);
        if (!policies.MigrationRequired)
        {
            return new PolicyMigrationResult(path, false, false, null, 1);
        }

        string content = PolicyDocument.SerializeVersioned(policies);
        if (!write)
        {
            return new PolicyMigrationResult(path, true, false, null, 1);
        }

        PolicyFile.EnsureMigrationWritable(path);
        PolicyFile.EnsureProtected(path, policies);

        string directory = Path.GetDirectoryName(path)
            ?? throw new IOException($"Policy path '{path}' has no parent directory.");
        string backup = path + $".legacy.{DateTime.UtcNow:yyyyMMddHHmmss}.{Guid.NewGuid():N}.bak";
        string temporary = Path.Combine(
            directory,
            $".{Path.GetFileName(path)}.migration-{Guid.NewGuid():N}.tmp");
        bool replaced = false;

        try
        {
            File.Copy(path, backup, overwrite: false);
            WriteAndFlush(temporary, content);
            PreserveUnixMode(path, temporary);

            if (OperatingSystem.IsWindows())
            {
                File.Replace(temporary, path, null, ignoreMetadataErrors: true);
            }
            else
            {
                File.Move(temporary, path, overwrite: true);
            }

            replaced = true;
            PolicyFile.EnsureProtected(path, policies);
            return new PolicyMigrationResult(path, true, true, backup, 1);
        }
        catch
        {
            if (replaced && File.Exists(backup))
            {
                File.Move(backup, path, overwrite: true);
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
