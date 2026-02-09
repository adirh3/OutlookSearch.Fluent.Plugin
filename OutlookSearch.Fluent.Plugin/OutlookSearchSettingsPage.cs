using Avalonia.Media;
using Avalonia.Media.Immutable;
using Blast.API.Settings;
using Blast.Core.Objects;

namespace OutlookSearch.Fluent.Plugin;

public sealed class OutlookSearchSettingsPage : SearchApplicationSettingsPage
{
    public OutlookSearchSettingsPage(SearchApplicationInfo searchApplicationInfo)
        : base(searchApplicationInfo)
    {
        SettingLevel = 2;
        IconGlyphColor = new ImmutableSolidColorBrush(Color.FromRgb(30, 144, 255));
    }

    // ── Email search ──

    [Setting(
        Name = nameof(SearchEmails),
        DisplayedName = "Search emails",
        Description = "Include email results when searching",
        SettingCategoryName = "Email",
        IconGlyph = "\uE715",
        DefaultValue = true,
        SettingLevel = 1)]
    public bool SearchEmails { get; set; } = true;

    [Setting(
        Name = nameof(MaxEmailResults),
        DisplayedName = "Maximum email results",
        Description = "Maximum number of email results per search",
        SettingCategoryName = "Email",
        IconGlyph = "\uE715",
        DefaultValue = 15,
        MinValue = 1,
        MaxValue = 50,
        SettingLevel = 2)]
    public int MaxEmailResults { get; set; } = 15;

    [Setting(
        Name = nameof(EmailSearchDaysBack),
        DisplayedName = "Email search range (days)",
        Description = "How many days back to search for emails. Larger values search more mail but may be slower.",
        SettingCategoryName = "Email",
        IconGlyph = "\uE8BF",
        DefaultValue = 90,
        MinValue = 7,
        MaxValue = 365,
        SettingLevel = 3)]
    public int EmailSearchDaysBack { get; set; } = 90;

    // ── Calendar search ──

    [Setting(
        Name = nameof(SearchEvents),
        DisplayedName = "Search calendar events",
        Description = "Include calendar event results when searching",
        SettingCategoryName = "Calendar",
        IconGlyph = "\uE787",
        DefaultValue = true,
        SettingLevel = 1)]
    public bool SearchEvents { get; set; } = true;

    [Setting(
        Name = nameof(MaxEventResults),
        DisplayedName = "Maximum event results",
        Description = "Maximum number of calendar event results per search",
        SettingCategoryName = "Calendar",
        IconGlyph = "\uE787",
        DefaultValue = 10,
        MinValue = 1,
        MaxValue = 30,
        SettingLevel = 2)]
    public int MaxEventResults { get; set; } = 10;

    [Setting(
        Name = nameof(EventSearchDaysBack),
        DisplayedName = "Past events range (days)",
        Description = "How many days back to search for past events",
        SettingCategoryName = "Calendar",
        IconGlyph = "\uE8BF",
        DefaultValue = 30,
        MinValue = 1,
        MaxValue = 180,
        SettingLevel = 3)]
    public int EventSearchDaysBack { get; set; } = 30;

    [Setting(
        Name = nameof(EventFutureDays),
        DisplayedName = "Future events range (days)",
        Description = "How many days ahead to search for upcoming events",
        SettingCategoryName = "Calendar",
        IconGlyph = "\uE8BF",
        DefaultValue = 90,
        MinValue = 7,
        MaxValue = 365,
        SettingLevel = 4)]
    public int EventFutureDays { get; set; } = 90;

    [Setting(
        Name = nameof(ShowUpcomingEventsOnEmpty),
        DisplayedName = "Show upcoming events on empty search",
        Description = "When using the outlook tag with empty search text, show your upcoming events",
        SettingCategoryName = "Calendar",
        IconGlyph = "\uE787",
        DefaultValue = true,
        SettingLevel = 5)]
    public bool ShowUpcomingEventsOnEmpty { get; set; } = true;

    // ── Quick actions ──

    [Setting(
        Name = nameof(ShowQuickActions),
        DisplayedName = "Show quick actions",
        Description = "Show quick action results like 'New Email' and 'New Meeting'",
        SettingCategoryName = "Quick Actions",
        IconGlyph = "\uE8FB",
        DefaultValue = true,
        SettingLevel = 1)]
    public bool ShowQuickActions { get; set; } = true;

    // ── Advanced ──

    [Setting(
        Name = nameof(GraphClientId),
        DisplayedName = "Microsoft Graph Client ID",
        Description = "Custom Azure AD app client ID for Graph API access. Leave empty to use the built-in default.",
        SettingCategoryName = "Advanced",
        IconGlyph = "\uE943",
        DefaultValue = "",
        IsAdvanced = true,
        SettingLevel = 1)]
    public string GraphClientId { get; set; } = string.Empty;
}
