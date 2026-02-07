using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Blast.API.Controllers.Speech;
using Blast.API.Core.Controllers.Keyboard;
using Blast.API.Core.Controllers.Speech;
using Blast.API.Core.Hotkeys;
using Blast.API.Core.Interfaces;
using Blast.API.Hotkeys;
using Blast.Core;
using Blast.Core.Interfaces;
using Blast.Core.Objects;
using Xunit;

namespace OutlookSearch.Fluent.Plugin.Tests;

/// <summary>
/// No-op implementations for static singletons required by SearchApplicationSettingsPage.
/// </summary>
file sealed class NoOpHotkeyManager : IHotkeyManager
{
    public void ToggleKeyHook(KeyHookType keyHookType, bool keyHook) { }
    public KeyState GetKeyState(VirtualKey virtualKey) => default;
    public bool CanRegisterHotkey(List<VirtualKey> hotkeys) => true;
    public bool IsHotkeyRegistered(Hotkey hotkey) => false;
    public bool RegisterHotkey(Hotkey hotkey) => true;
    public bool UnregisterHotkey(Hotkey hotkey) => true;
    public event EventHandler<Hotkey>? HotkeyPressed;
    public event EventHandler<KeyPressedEventArgs>? KeyPressed;
    public event EventHandler<KeyPressedEventArgs>? KeyPressedSynchronous;
    public event EventHandler<MousePressedEventArgs>? MousePressed;
    public void Dispose() { }
}

file sealed class NoOpSpeechController : ISpeechController
{
    public bool UseSFX => false;
    public bool IsEnabled => false;
    public void RepeatLastSpeak() { }
    public void QueueSpeak(string text, bool cancelPrevious = false, bool waitForSpeech = false, bool allowCancel = true) { }
    public void CancelAll() { }
    public void DescribeSearchResult(ISearchResult searchResult, ICollection<ISearchResult> searchResults) { }
    public void DescribeSearchOperation(ISearchOperation searchOperation, bool cancelPrevious = true, bool waitForSpeech = false, bool speakKeyGesture = true) { }
    public event EventHandler<bool>? OnSpeechModeToggled;
}

public sealed class PluginContractTests
{
    static PluginContractTests()
    {
        // Initialize static singletons required by SearchApplicationSettingsPage
        if (HotkeyManager.HotkeyManagerInstance == null)
            HotkeyManager.HotkeyManagerInstance = new NoOpHotkeyManager();
        if (SpeechControllerUtils.SpeechController == null)
            SpeechControllerUtils.SpeechController = new NoOpSpeechController();
    }

    [Fact]
    public void SearchApp_implements_ISearchApplication_with_parameterless_ctor()
    {
        Type appType = typeof(OutlookSearchSearchApp);
        Assert.True(typeof(ISearchApplication).IsAssignableFrom(appType));
        Assert.NotNull(appType.GetConstructor(Type.EmptyTypes));
    }

    [Fact]
    public void GetApplicationInfo_returns_valid_metadata()
    {
        var app = new OutlookSearchSearchApp();
        SearchApplicationInfo info = app.GetApplicationInfo();

        Assert.NotNull(info);
        Assert.Equal("Outlook Search", info.Name);
        Assert.NotNull(info.DefaultSearchTags);
        Assert.True(info.DefaultSearchTags.Count >= 3, "Should have outlook, email, and calendar tags");
        Assert.True(info.SearchTagOnly, "Plugin should be search-tag only");
        Assert.False(info.IsProcessSearchEnabled);
    }

    [Fact]
    public void GetApplicationInfo_has_all_expected_tags()
    {
        var app = new OutlookSearchSearchApp();
        SearchApplicationInfo info = app.GetApplicationInfo();

        var tagNames = info.DefaultSearchTags.Select(t => t.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
        Assert.Contains("outlook", tagNames);
        Assert.Contains("email", tagNames);
        Assert.Contains("calendar", tagNames);
    }

    [Fact]
    public async Task SearchAsync_returns_quick_actions_for_empty_search_with_outlook_tag()
    {
        var app = new OutlookSearchSearchApp();
        var request = new SearchRequest("", "outlook", SearchType.SearchAll);

        int count = 0;
        await foreach (ISearchResult result in app.SearchAsync(request, CancellationToken.None))
        {
            Assert.NotNull(result);
            count++;
        }

        // Should return quick actions + upcoming events (if Outlook connected) or just quick actions
        Assert.True(count > 0, "Should return at least quick action results for empty outlook tag search");
    }

    [Fact]
    public async Task SearchAsync_returns_no_results_for_unrelated_tag()
    {
        var app = new OutlookSearchSearchApp();
        var request = new SearchRequest("smoke", "not-the-plugin-tag", SearchType.SearchAll);

        bool yielded = false;
        await foreach (ISearchResult _ in app.SearchAsync(request, CancellationToken.None))
        {
            yielded = true;
            break;
        }

        Assert.False(yielded);
    }

    [Fact]
    public async Task SearchAsync_respects_cancellation()
    {
        var app = new OutlookSearchSearchApp();
        using var cts = new CancellationTokenSource();
        await cts.CancelAsync();

        var request = new SearchRequest("test", "outlook", SearchType.SearchAll);

        bool yielded = false;
        await foreach (ISearchResult _ in app.SearchAsync(request, cts.Token))
        {
            yielded = true;
            break;
        }

        Assert.False(yielded, "Should not return results when cancelled");
    }

    [Fact]
    public async Task SearchAsync_skips_process_search()
    {
        var app = new OutlookSearchSearchApp();
        var request = new SearchRequest("test", "outlook", SearchType.SearchProcess);

        bool yielded = false;
        await foreach (ISearchResult _ in app.SearchAsync(request, CancellationToken.None))
        {
            yielded = true;
            break;
        }

        Assert.False(yielded, "Should not support process search");
    }

    [Fact]
    public void SearchApp_wires_built_in_settings_page()
    {
        var app = new OutlookSearchSearchApp();
        SearchApplicationInfo appInfo = app.GetApplicationInfo();

        Assert.NotNull(appInfo.SettingsPage);
        Assert.IsType<OutlookSearchSettingsPage>(appInfo.SettingsPage);
    }

    [Fact]
    public void Settings_have_reasonable_defaults()
    {
        var app = new OutlookSearchSearchApp();
        var settings = (OutlookSearchSettingsPage)app.GetApplicationInfo().SettingsPage;
        Assert.Equal(15, settings.MaxEmailResults);
        Assert.Equal(10, settings.MaxEventResults);
        Assert.Equal(90, settings.EmailSearchDaysBack);
        Assert.Equal(30, settings.EventSearchDaysBack);
        Assert.Equal(90, settings.EventFutureDays);
        Assert.True(settings.ShowQuickActions);
        Assert.True(settings.SearchEmails);
        Assert.True(settings.SearchEvents);
        Assert.True(settings.ShowUpcomingEventsOnEmpty);
        Assert.Equal(string.Empty, settings.GraphClientId);
    }

    [Fact]
    public void QuickAction_results_have_correct_structure()
    {
        var result = new OutlookActionSearchResult(
            "Compose New Email",
            "Open a blank new email window in Outlook",
            "\uE70F",
            "compose",
            new System.Collections.ObjectModel.ObservableCollection<ISearchOperation>(),
            new System.Collections.ObjectModel.ObservableCollection<Blast.Core.Results.SearchTag>(),
            3.0);

        Assert.Equal("Compose New Email", result.DisplayedName);
        Assert.Equal("compose", result.ActionId);
        Assert.Equal("Action", result.MLResultType);
        Assert.Equal("Quick Actions", result.GroupName);
        Assert.NotNull(result.InformationElements);
    }

    [Fact]
    public void Email_search_result_formats_correctly()
    {
        var email = new OutlookEmailItem
        {
            EntryId = "test-entry-id",
            Subject = "Meeting Notes",
            SenderName = "John Doe",
            SenderEmail = "john@example.com",
            ReceivedTime = DateTime.Now.AddHours(-2),
            BodyPreview = "Here are the notes from today...",
            HasAttachments = true,
            IsRead = false,
            ToRecipients = "me@example.com",
            FolderName = "Inbox",
            Importance = 2
        };

        var result = new OutlookEmailSearchResult(
            email,
            new System.Collections.ObjectModel.ObservableCollection<ISearchOperation>(),
            new System.Collections.ObjectModel.ObservableCollection<Blast.Core.Results.SearchTag>(),
            8.0);

        Assert.Equal("Meeting Notes", result.DisplayedName);
        Assert.Equal("Emails", result.GroupName);
        Assert.Contains("John Doe", result.AdditionalInformation);
        Assert.NotNull(result.InformationElements);
        Assert.True(result.InformationElements.Count >= 4);
    }

    [Fact]
    public void Event_search_result_formats_correctly()
    {
        var calendarItem = new OutlookCalendarItem
        {
            EntryId = "test-event-id",
            Subject = "Team Standup",
            StartTime = DateTime.Now.AddMinutes(30),
            EndTime = DateTime.Now.AddMinutes(60),
            Location = "Room 42",
            Organizer = "Jane Smith",
            RequiredAttendees = "john@example.com; bob@example.com"
        };

        var result = new OutlookEventSearchResult(
            calendarItem,
            new System.Collections.ObjectModel.ObservableCollection<ISearchOperation>(),
            new System.Collections.ObjectModel.ObservableCollection<Blast.Core.Results.SearchTag>(),
            6.0);

        Assert.Equal("Team Standup", result.DisplayedName);
        Assert.Equal("Events", result.GroupName);
        Assert.Contains("Room 42", result.AdditionalInformation);
        Assert.NotNull(result.InformationElements);
        Assert.True(result.InformationElements.Count >= 2);
    }
}
