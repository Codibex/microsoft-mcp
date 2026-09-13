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
    public void Guard_blocks_external_recipients()
    {
        var ex = Assert.Throws<MailServiceException>(() =>
            RecipientGuard.ValidateRecipients(["a@firma.de", "b@gmail.com"], InternalPolicy()));

        Assert.Contains("[invalid-request]", ex.Message);
    }

    [Fact]
    public void Guard_rejects_malformed_addresses()
    {
        Assert.Throws<MailServiceException>(() =>
            RecipientGuard.ValidateRecipients(["keine-domain"], InternalPolicy()));
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
    [InlineData("Text<br/>weiter", true)]
    [InlineData("Hallo Welt", false)]
    [InlineData("a < b und c > d", false)]
    [InlineData("2 <3 Äpfel", false)]
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
                  "aiDisclosureEnabled": true,
                  "aiDisclosureText": "KI-Hinweis.",
                }
                """);

            var policy = PolicyFile.Load(path);

            Assert.True(policy.RequireInternalRecipients);
            Assert.Single(policy.AllowedRecipientDomains);
            Assert.Equal("firma.de", policy.AllowedRecipientDomains[0]);
            Assert.Equal("KI-Hinweis.", policy.AiDisclosureText);
        }
        finally
        {
            File.Delete(path);
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
    public void EnsureProtected_passes_for_readonly_restrictive_file()
    {
        if (IsElevatedTestProcess())
        {
            return; // Root/admin bypasses permission bits; nothing meaningful to assert.
        }

        string path = Path.GetTempFileName();
        try
        {
            File.WriteAllText(path, "{}");
            MakeReadOnly(path);
            Assert.False(PolicyFile.IsWritable(path));

            PolicyFile.EnsureProtected(path, InternalPolicy());
        }
        finally
        {
            MakeWritable(path);
            File.Delete(path);
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

    private static bool IsElevatedTestProcess() =>
        string.Equals(Environment.UserName, "root", StringComparison.OrdinalIgnoreCase);
}
