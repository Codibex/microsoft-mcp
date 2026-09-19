using MicrosoftMcp.Common;

namespace MicrosoftMcp.Common.Tests;

public sealed class PolicyDocumentTests
{
    [Fact]
    public void Legacy_policy_applies_restrictions_to_outlook_and_calendar()
    {
        EffectivePolicySet policies = PolicyDocument.Parse("""
            {
              "requireInternalRecipients": true,
              "allowedRecipientDomains": ["firma.de"],
              "allowedRecipientAddresses": ["owner@partner.example"],
              "aiDisclosureEnabled": true,
              "aiDisclosureText": "Hinweis"
            }
            """);

        Assert.True(policies.IsLegacy);
        Assert.True(policies.MigrationRequired);
        Assert.True(policies.Outlook.RequireInternalRecipients);
        Assert.True(policies.Outlook.AiDisclosureEnabled);
        Assert.True(policies.Calendar.RequireInternalAttendees);
        Assert.Equal(["firma.de"], policies.Calendar.AllowedAttendeeDomains);
        Assert.Empty(policies.Teams.GetType().GetProperties());
    }

    [Fact]
    public void Explicit_sections_override_only_their_service()
    {
        EffectivePolicySet policies = PolicyDocument.Parse("""
            {
              "version": 1,
              "requireInternalRecipients": true,
              "allowedRecipientDomains": ["legacy.example"],
              "aiDisclosureEnabled": true,
              "aiDisclosureText": "Legacy",
              "outlook": {
                "requireInternalRecipients": false,
                "allowedRecipientDomains": [],
                "allowedRecipientAddresses": [],
                "aiDisclosureEnabled": false,
                "aiDisclosureText": ""
              },
              "calendar": {}
            }
            """);

        Assert.True(policies.MigrationRequired);
        Assert.False(policies.Outlook.RequireInternalRecipients);
        Assert.False(policies.Outlook.AiDisclosureEnabled);
        Assert.False(policies.Calendar.RequireInternalAttendees);
        Assert.Empty(policies.Calendar.AllowedAttendeeDomains);
    }

    [Fact]
    public void Calendar_does_not_accept_outlook_disclosure_fields()
    {
        Assert.Throws<System.Text.Json.JsonException>(() => PolicyDocument.Parse("""
            {
              "version": 1,
              "calendar": {
                "aiDisclosureEnabled": true
              }
            }
            """));
    }

    [Fact]
    public void Service_sections_require_a_supported_version()
    {
        Assert.Throws<System.Text.Json.JsonException>(() => PolicyDocument.Parse("""
            {
              "calendar": {}
            }
            """));

        Assert.Throws<System.Text.Json.JsonException>(() => PolicyDocument.Parse("""
            {
              "version": 2,
              "calendar": {}
            }
            """));
    }

    [Fact]
    public void Migration_preview_and_write_are_idempotent()
    {
        string path = Path.Combine(Path.GetTempPath(), $"policy-{Guid.NewGuid():N}.json");
        try
        {
            File.WriteAllText(path, """{"requireInternalRecipients":false}""");

            PolicyMigrationResult preview = PolicyMigrator.Migrate(path, write: false);
            Assert.True(preview.MigrationRequired);
            Assert.False(preview.Written);
            Assert.Contains("requireInternalRecipients", File.ReadAllText(path));

            PolicyMigrationResult written = PolicyMigrator.Migrate(path, write: true);
            Assert.True(written.Written);
            Assert.NotNull(written.BackupPath);
            Assert.True(File.Exists(written.BackupPath));
            Assert.False(PolicyDocument.Load(path).MigrationRequired);

            PolicyMigrationResult second = PolicyMigrator.Migrate(path, write: true);
            Assert.False(second.MigrationRequired);
            Assert.False(second.Written);
        }
        finally
        {
            File.Delete(path);
            foreach (string backup in Directory.GetFiles(Path.GetDirectoryName(path)!, Path.GetFileName(path) + ".legacy.*.bak"))
            {
                File.Delete(backup);
            }
        }
    }

      [Fact]
      public void Failed_migration_does_not_create_backup_or_change_source()
      {
        string path = Path.Combine(Path.GetTempPath(), $"policy-invalid-{Guid.NewGuid():N}.json");
        const string source = "{\"unknown\":true}";
        try
        {
          File.WriteAllText(path, source);

          Assert.Throws<System.Text.Json.JsonException>(() => PolicyMigrator.Migrate(path, write: true));
          Assert.Equal(source, File.ReadAllText(path));
          Assert.Empty(Directory.GetFiles(
            Path.GetDirectoryName(path)!,
            Path.GetFileName(path) + ".legacy.*.bak"));
        }
        finally
        {
          File.Delete(path);
        }
      }
}
