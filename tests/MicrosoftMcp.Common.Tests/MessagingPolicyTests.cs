namespace MicrosoftMcp.Common.Tests;

public sealed class MessagingPolicyTests
{
    private static MessagingPolicyOptions InternalPolicy() => new()
    {
        RequireInternalRecipients = true,
        AllowedRecipientDomains = ["firma.de", "tochter.firma.de"],
        AiDisclosureEnabled = true,
        AiDisclosureText = "KI-Hinweis: Entwurf prüfen."
    };

    // Validator

    [Fact]
    public void Validator_accepts_restrictive_policy()
    {
        Assert.True(new MessagingPolicyValidator().Validate(null, InternalPolicy()).Succeeded);
    }

    [Fact]
    public void Validator_accepts_permissive_defaults()
    {
        Assert.True(new MessagingPolicyValidator().Validate(null, new MessagingPolicyOptions()).Succeeded);
    }

    [Fact]
    public void Validator_rejects_internal_without_domains()
    {
        var result = new MessagingPolicyValidator().Validate(null, new MessagingPolicyOptions
        {
            RequireInternalRecipients = true
        });

        Assert.False(result.Succeeded);
    }

    [Fact]
    public void Validator_accepts_internal_with_exact_addresses_only()
    {
        var result = new MessagingPolicyValidator().Validate(null, new MessagingPolicyOptions
        {
            RequireInternalRecipients = true,
            AllowedRecipientAddresses = ["owner@example.com"]
        });

        Assert.True(result.Succeeded);
    }

    [Fact]
    public void Validator_rejects_disclosure_without_text()
    {
        var result = new MessagingPolicyValidator().Validate(null, new MessagingPolicyOptions
        {
            AiDisclosureEnabled = true
        });

        Assert.False(result.Succeeded);
    }

    // RecipientGuard

    [Fact]
    public void Guard_allows_internal_recipients()
    {
        var policy = InternalPolicy();
        RecipientGuard.ValidateRecipients(["a@firma.de", "B@Tochter.Firma.DE"], policy);
    }

    [Fact]
    public void Guard_allows_subdomains()
    {
        RecipientGuard.ValidateRecipients(["a@mail.firma.de"], InternalPolicy());
    }

    [Fact]
    public void Guard_allows_exact_address_exception_to_domain_policy()
    {
        var policy = new MessagingPolicyOptions
        {
            RequireInternalRecipients = true,
            AllowedRecipientAddresses = ["Guest@partner.example"]
        };

        RecipientGuard.ValidateRecipients(["guest@partner.example"], policy);
    }

    [Fact]
    public void Guard_blocks_external_recipients()
    {
        var ex = Assert.Throws<GraphServiceException>(() =>
            RecipientGuard.ValidateRecipients(["a@firma.de", "b@gmail.com"], InternalPolicy()));

        Assert.Contains("[invalid-request]", ex.Message);
    }

    [Fact]
    public void Guard_rejects_malformed_addresses()
    {
        Assert.Throws<GraphServiceException>(() =>
            RecipientGuard.ValidateRecipients(["keine-domain"], InternalPolicy()));
    }

    [Fact]
    public void Guard_rejects_double_at_smuggling()
    {
        // LastIndexOf would read firma.de and pass; exactly one '@' is required.
        var ex = Assert.Throws<GraphServiceException>(() =>
            RecipientGuard.ValidateRecipients(["a@extern.example@firma.de"], InternalPolicy()));

        Assert.Contains("[invalid-request]", ex.Message);
    }

    [Fact]
    public void Guard_disabled_allows_everything()
    {
        RecipientGuard.ValidateRecipients(["a@gmail.com"], new MessagingPolicyOptions());
    }

    // MessageDisclosure

    [Fact]
    public void Disclosure_appends_text_body()
    {
        string result = MessageDisclosure.Apply("Hallo", isHtml: false, InternalPolicy());

        Assert.Contains("Hallo", result);
        Assert.Contains("KI-Hinweis", result);
    }

    [Fact]
    public void Disclosure_appends_html_body_encoded()
    {
        var policy = InternalPolicy();
        policy.AiDisclosureText = "Prüfen & freigeben <danke>";

        string result = MessageDisclosure.Apply("<p>Hallo</p>", isHtml: true, policy);

        Assert.Contains("<hr>", result);
        Assert.Contains("Pr&#252;fen &amp; freigeben &lt;danke&gt;", result);
        Assert.DoesNotContain("<danke>", result);
    }

    [Fact]
    public void Disclosure_is_idempotent()
    {
        var policy = InternalPolicy();
        string once = MessageDisclosure.Apply("Hallo", isHtml: false, policy);

        Assert.Equal(once, MessageDisclosure.Apply(once, isHtml: false, policy));
    }

    [Fact]
    public void Disclosure_disabled_returns_body_unchanged()
    {
        Assert.Equal("Hallo", MessageDisclosure.Apply("Hallo", isHtml: false, new MessagingPolicyOptions()));
    }

    [Fact]
    public void Disclosure_empty_body_yields_text_only()
    {
        Assert.Equal("KI-Hinweis: Entwurf prüfen.", MessageDisclosure.Apply(string.Empty, isHtml: false, InternalPolicy()));
    }

    [Theory]
    [InlineData("<p>Hallo</p>", true)]
    [InlineData("<p>", true)]
    [InlineData("<br>", true)]
    [InlineData("Text<br/>weiter", true)]
    [InlineData("Hallo Welt", false)]
    [InlineData("a < b und c > d", false)]
    [InlineData("2 <3 Äpfel", false)]
    [InlineData("<>", false)]
    public void LooksLikeHtml_detects_tags_not_comparisons(string body, bool expected)
    {
        Assert.Equal(expected, MessageDisclosure.LooksLikeHtml(body));
    }

    [Fact]
    public void ApplyAuto_uses_html_variant_for_html_comments()
    {
        string result = MessageDisclosure.ApplyAuto("<p>Bin dabei</p>", InternalPolicy());

        Assert.Contains("<hr>", result);
        Assert.DoesNotContain("\n\n---\n", result);
    }

    [Fact]
    public void ApplyAuto_uses_text_variant_for_text_comments()
    {
        string result = MessageDisclosure.ApplyAuto("Bin dabei", InternalPolicy());

        Assert.Contains("\n\n---\n", result);
        Assert.DoesNotContain("<hr>", result);
    }

    // PolicyFile

    [Fact]
    public void Load_parses_json_with_comments_and_trailing_commas()
    {
        string path = Path.GetTempFileName();
        try
        {
            File.WriteAllText(path, """
                {
                  // Admin-owned policy
                  "requireInternalRecipients": true,
                  "allowedRecipientDomains": [ "firma.de", ],
                  "allowedRecipientAddresses": [ "owner@partner.example", ],
                  "aiDisclosureEnabled": true,
                  "aiDisclosureText": "KI-Hinweis.",
                }
                """);

            var policy = PolicyFile.Load(path);

            Assert.True(policy.RequireInternalRecipients);
            Assert.Single(policy.AllowedRecipientDomains);
            Assert.Equal("firma.de", policy.AllowedRecipientDomains[0]);
            Assert.Single(policy.AllowedRecipientAddresses);
            Assert.Equal("owner@partner.example", policy.AllowedRecipientAddresses[0]);
            Assert.Equal("KI-Hinweis.", policy.AiDisclosureText);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Load_normalizes_null_domains_to_empty()
    {
        string path = Path.GetTempFileName();
        try
        {
            File.WriteAllText(path, """{"requireInternalRecipients": false, "allowedRecipientDomains": null, "allowedRecipientAddresses": null}""");

            var policy = PolicyFile.Load(path);

            Assert.NotNull(policy.AllowedRecipientDomains);
            Assert.Empty(policy.AllowedRecipientDomains);
            Assert.NotNull(policy.AllowedRecipientAddresses);
            Assert.Empty(policy.AllowedRecipientAddresses);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Load_rejects_unknown_properties_fail_closed()
    {
        string path = Path.GetTempFileName();
        try
        {
            // Typo: must not silently start with a permissive policy.
            File.WriteAllText(path, """{"requireInternalRecipient": true}""");

            Assert.Throws<System.Text.Json.JsonException>(() => PolicyFile.Load(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Load_rejects_null_document()
    {
        string path = Path.GetTempFileName();
        try
        {
            File.WriteAllText(path, "null");

            Assert.Throws<System.Text.Json.JsonException>(() => PolicyFile.Load(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ProbeExists_distinguishes_missing_from_unreadable()
    {
        Assert.False(PolicyFile.ProbeExists(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"))));

        if (OperatingSystem.IsWindows() || IsElevatedTestProcess())
        {
            return; // Needs Unix permission bits; root bypasses them anyway.
        }

        // Inaccessible directory: stat fails (Exists=false) but opening would be
        // EACCES — must fail closed instead of falling back to a weaker policy.
        string dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            File.SetUnixFileMode(dir, UnixFileMode.None);

            Assert.Throws<Microsoft.Extensions.Options.OptionsValidationException>(() =>
                PolicyFile.ProbeExists(Path.Combine(dir, "policy.json")));
        }
        finally
        {
            File.SetUnixFileMode(dir,
                UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
            Directory.Delete(dir);
        }
    }

    [Fact]
    public void EnsureProtected_passes_for_permissive_policy_on_writable_file()
    {
        string path = Path.GetTempFileName();
        try
        {
            // Permissive default policy: nothing enforced, nothing to protect.
            PolicyFile.EnsureProtected(path, new MessagingPolicyOptions());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void EnsureProtected_passes_for_readonly_file_in_readonly_dir()
    {
        if (IsElevatedTestProcess())
        {
            return; // Root/admin bypasses permission bits; nothing meaningful to assert.
        }

        if (OperatingSystem.IsWindows())
        {
            return; // Read-only dirs need ACL juggling; covered by file-level test below.
        }

        string dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        string path = Path.Combine(dir, "policy.json");
        try
        {
            File.WriteAllText(path, "{}");
            MakeReadOnly(path);
            Assert.False(PolicyFile.IsWritable(path));
            File.SetUnixFileMode(dir,
                UnixFileMode.UserRead | UnixFileMode.UserExecute |
                UnixFileMode.GroupRead | UnixFileMode.GroupExecute |
                UnixFileMode.OtherRead | UnixFileMode.OtherExecute);
            Assert.False(PolicyFile.IsDirectoryWritable(dir));

            PolicyFile.EnsureProtected(path, InternalPolicy());
        }
        finally
        {
            File.SetUnixFileMode(dir,
                UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
            MakeWritable(path);
            File.Delete(path);
            Directory.Delete(dir);
        }
    }

    [Fact]
    public void EnsureProtected_throws_for_readonly_file_in_writable_dir()
    {
        // Delete-and-replace defeats file-level read-only: the directory counts too.
        if (IsElevatedTestProcess())
        {
            return;
        }

        string dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        string path = Path.Combine(dir, "policy.json");
        try
        {
            File.WriteAllText(path, "{}");
            MakeReadOnly(path);
            Assert.True(PolicyFile.IsDirectoryWritable(dir));

            Assert.Throws<Microsoft.Extensions.Options.OptionsValidationException>(() =>
                PolicyFile.EnsureProtected(path, InternalPolicy()));
        }
        finally
        {
            MakeWritable(path);
            File.Delete(path);
            Directory.Delete(dir);
        }
    }

    [Fact]
    public void EnsureProtected_throws_for_writable_restrictive_file()
    {
        if (IsElevatedTestProcess())
        {
            return; // Elevated processes skip the check by design (admin testing).
        }

        string path = Path.GetTempFileName();
        try
        {
            File.WriteAllText(path, "{}");
            Assert.True(PolicyFile.IsWritable(path));

            Assert.Throws<Microsoft.Extensions.Options.OptionsValidationException>(() =>
                PolicyFile.EnsureProtected(path, InternalPolicy()));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SystemPolicyPath_is_os_specific()
    {
        string path = PolicyFile.SystemPolicyPath;

        if (OperatingSystem.IsWindows())
        {
            Assert.Contains("microsoft-mcp", path);
            Assert.EndsWith("policy.json", path);
        }
        else if (OperatingSystem.IsMacOS())
        {
            Assert.StartsWith("/Library/Application Support/", path);
        }
        else
        {
            Assert.Equal("/etc/microsoft-mcp/policy.json", path);
        }
    }

    private static void MakeReadOnly(string path)
    {
        if (OperatingSystem.IsWindows())
        {
            File.SetAttributes(path, File.GetAttributes(path) | FileAttributes.ReadOnly);
            return;
        }

        File.SetUnixFileMode(path,
            UnixFileMode.UserRead | UnixFileMode.GroupRead | UnixFileMode.OtherRead);
    }

    private static void MakeWritable(string path)
    {
        if (OperatingSystem.IsWindows())
        {
            File.SetAttributes(path, File.GetAttributes(path) & ~FileAttributes.ReadOnly);
            return;
        }

        File.SetUnixFileMode(path,
            UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.GroupRead | UnixFileMode.OtherRead);
    }

    private static bool IsElevatedTestProcess()
    {
        // Mirror PolicyFile.IsElevated: Windows administrator token counts too,
        // otherwise the Throws-tests fail when the suite runs elevated on Windows.
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
