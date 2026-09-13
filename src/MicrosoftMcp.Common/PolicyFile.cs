using System.Security.Cryptography;
using System.Text.Json;
using Microsoft.Extensions.Options;

namespace MicrosoftMcp.Common;

/// <summary>Admin-owned policy file handling, shared by all messaging servers.
/// Cross-platform (Windows, Linux, macOS).
///
/// Lookup order: OS-specific system location first, then <c>policy.json</c> next to
/// the server binary. The system locations require admin rights to write, which is
/// what keeps a user-level LLM (same OS user, file access via the agent harness)
/// from rewriting the policy:
/// <list type="bullet">
/// <item>Windows: <c>%ProgramData%\microsoft-mcp\policy.json</c> (Users = read)</item>
/// <item>macOS: <c>/Library/Application Support/microsoft-mcp/policy.json</c> (root-owned)</item>
/// <item>Linux: <c>/etc/microsoft-mcp/policy.json</c> (root-owned)</item>
/// </list></summary>
public static class PolicyFile
{
    public const string FileName = "policy.json";

    /// <summary>OS-specific system location for the admin-owned policy file.</summary>
    public static string SystemPolicyPath =>
        OperatingSystem.IsWindows()
            ? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                "microsoft-mcp", FileName)
            : OperatingSystem.IsMacOS()
                ? Path.Combine("/Library/Application Support", "microsoft-mcp", FileName)
                : Path.Combine("/etc", "microsoft-mcp", FileName);

    /// <summary>First existing file wins: system path, then next to the binary.</summary>
    public static string? FindPolicyFile()
    {
        string system = SystemPolicyPath;
        if (File.Exists(system))
        {
            return system;
        }

        string local = Path.Combine(AppContext.BaseDirectory, FileName);
        return File.Exists(local) ? local : null;
    }

    /// <summary>Parses policy.json (comments and trailing commas allowed).
    /// Never reads env vars or user-secrets: this file is the only source.</summary>
    public static MessagingPolicyOptions Load(string path)
    {
        string json = File.ReadAllText(path);
        var options = JsonSerializer.Deserialize<MessagingPolicyOptions>(json, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            ReadCommentHandling = JsonCommentHandling.Skip,
            AllowTrailingCommas = true
        }) ?? new MessagingPolicyOptions();
        options.AllowedRecipientDomains = [.. options.AllowedRecipientDomains.Where(d => !string.IsNullOrWhiteSpace(d))];
        return options;
    }

    /// <summary>SHA-256 of the file bytes (hex), for startup-log audit on stderr.</summary>
    public static string ComputeHash(string path)
    {
        using var sha = SHA256.Create();
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(sha.ComputeHash(stream)).ToLowerInvariant();
    }

    /// <summary>Fails fast when a restrictive policy file is writable by the current
    /// non-elevated process. A root/admin-owned read-only file (standard deployment)
    /// is not writable and passes on all three OSes. Permissive policies (nothing
    /// enforced) skip the check — there is nothing to protect.</summary>
    public static void EnsureProtected(string path, MessagingPolicyOptions policy)
    {
        if (!policy.IsRestrictive)
        {
            return;
        }

        if (IsElevated())
        {
            Console.Error.WriteLine(
                $"[startup] Policy file '{path}' is writable, but the process runs elevated – " +
                "protection check skipped (admin testing scenario).");
            return;
        }

        if (IsWritable(path))
        {
            throw new OptionsValidationException(
                "Messaging",
                typeof(MessagingPolicyOptions),
                [$"Policy file '{path}' enforces restrictions but is writable by the current user, " +
                 "so it could be rewritten. Next: deploy it admin-owned and read-only " +
                 "(Linux/macOS: chown root, chmod 644; Windows: admin-owned file under %ProgramData% " +
                 "with Users=read), see docs/outlook.md."]);
        }
    }

    /// <summary>Probes whether the current process can open the file for writing
    /// (without modifying it). Used by <see cref="EnsureProtected"/>.</summary>
    public static bool IsWritable(string path)
    {
        try
        {
            using var _ = new FileStream(path, FileMode.Open, FileAccess.Write, FileShare.ReadWrite);
            return true;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
        catch (IOException)
        {
            // Sharing violation, locked volume, etc.: not provably writable.
            return false;
        }
        catch (NotSupportedException)
        {
            return false;
        }
    }

    private static bool IsElevated()
    {
        if (OperatingSystem.IsWindows())
        {
            try
            {
                using var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
                var principal = new System.Security.Principal.WindowsPrincipal(identity);
                return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
            }
            catch (Exception ex) when (ex is PlatformNotSupportedException or UnauthorizedAccessException)
            {
                return false;
            }
        }

        return string.Equals(Environment.UserName, "root", StringComparison.OrdinalIgnoreCase);
    }
}
