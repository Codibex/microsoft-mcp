using System.Text.Json;
using System.Text.Json.Serialization;

namespace MicrosoftMcp.Common;

/// <summary>Versioned policy document parsing and legacy expansion.</summary>
public static class PolicyDocument
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow
    };

    private static readonly JsonSerializerOptions WriteOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    public static EffectivePolicySet Load(string path)
    {
        return Parse(File.ReadAllText(path));
    }

    public static EffectivePolicySet Parse(string json)
    {
        using JsonDocument document = JsonDocument.Parse(json, new JsonDocumentOptions
        {
            CommentHandling = JsonCommentHandling.Skip,
            AllowTrailingCommas = true
        });

        if (document.RootElement.ValueKind != JsonValueKind.Object)
        {
            throw new JsonException("Policy document must contain an object.");
        }

        PolicyJsonModel model = JsonSerializer.Deserialize<PolicyJsonModel>(json, JsonOptions)
            ?? throw new JsonException("Policy document contains no policy object.");

        bool hasOutlook = HasProperty(document.RootElement, nameof(PolicyJsonModel.Outlook));
        bool hasCalendar = HasProperty(document.RootElement, nameof(PolicyJsonModel.Calendar));
        bool hasTeams = HasProperty(document.RootElement, nameof(PolicyJsonModel.Teams));
        bool hasServiceSection = hasOutlook || hasCalendar || hasTeams;
        bool hasLegacy = HasLegacyProperty(document.RootElement);

        if (model.Version is not null && model.Version != 1)
        {
            throw new JsonException($"Unsupported policy version '{model.Version}'. Supported version: 1.");
        }

        if (hasServiceSection && model.Version is null)
        {
            throw new JsonException("A version is required when policy service sections are present.");
        }

        if (hasOutlook && model.Outlook is null
            || hasCalendar && model.Calendar is null
            || hasTeams && model.Teams is null)
        {
            throw new JsonException("Policy service sections must be objects, not null.");
        }

        OutlookPolicyOptions legacyOutlook = new()
        {
            RequireInternalRecipients = model.RequireInternalRecipients ?? false,
            AllowedRecipientDomains = Normalize(model.AllowedRecipientDomains),
            AllowedRecipientAddresses = Normalize(model.AllowedRecipientAddresses),
            AiDisclosureEnabled = model.AiDisclosureEnabled ?? false,
            AiDisclosureText = model.AiDisclosureText ?? string.Empty
        };

        OutlookPolicyOptions outlook = hasOutlook
            ? Normalize(model.Outlook!)
            : legacyOutlook;

        CalendarPolicyOptions calendar = hasCalendar
            ? Normalize(model.Calendar!)
            : new CalendarPolicyOptions
            {
                RequireInternalAttendees = legacyOutlook.RequireInternalRecipients,
                AllowedAttendeeDomains = [.. legacyOutlook.AllowedRecipientDomains],
                AllowedAttendeeAddresses = [.. legacyOutlook.AllowedRecipientAddresses]
            };

        return new EffectivePolicySet(
            outlook,
            calendar,
            hasTeams ? model.Teams! : new TeamsPolicyOptions(),
            model.Version,
            model.Version is null || hasLegacy);
    }

    public static string SerializeVersioned(EffectivePolicySet policies)
    {
        return JsonSerializer.Serialize(new
        {
            version = 1,
            outlook = policies.Outlook,
            calendar = policies.Calendar,
            teams = policies.Teams
        }, WriteOptions) + Environment.NewLine;
    }

    private static OutlookPolicyOptions Normalize(OutlookPolicyOptions options)
    {
        options.AllowedRecipientDomains = Normalize(options.AllowedRecipientDomains);
        options.AllowedRecipientAddresses = Normalize(options.AllowedRecipientAddresses);
        return options;
    }

    private static CalendarPolicyOptions Normalize(CalendarPolicyOptions options)
    {
        options.AllowedAttendeeDomains = Normalize(options.AllowedAttendeeDomains);
        options.AllowedAttendeeAddresses = Normalize(options.AllowedAttendeeAddresses);
        return options;
    }

    private static string[] Normalize(string[]? values) =>
        [.. (values ?? []).Where(value => !string.IsNullOrWhiteSpace(value))];

    private static bool HasLegacyProperty(JsonElement root) =>
        HasProperty(root, nameof(PolicyJsonModel.RequireInternalRecipients))
        || HasProperty(root, nameof(PolicyJsonModel.AllowedRecipientDomains))
        || HasProperty(root, nameof(PolicyJsonModel.AllowedRecipientAddresses))
        || HasProperty(root, nameof(PolicyJsonModel.AiDisclosureEnabled))
        || HasProperty(root, nameof(PolicyJsonModel.AiDisclosureText));

    private static bool HasProperty(JsonElement root, string name) =>
        root.EnumerateObject().Any(property =>
            string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase));

    private sealed class PolicyJsonModel
    {
        public int? Version { get; set; }

        public OutlookPolicyOptions? Outlook { get; set; }

        public CalendarPolicyOptions? Calendar { get; set; }

        public TeamsPolicyOptions? Teams { get; set; }

        public bool? RequireInternalRecipients { get; set; }

        public string[]? AllowedRecipientDomains { get; set; }

        public string[]? AllowedRecipientAddresses { get; set; }

        public bool? AiDisclosureEnabled { get; set; }

        public string? AiDisclosureText { get; set; }
    }
}
