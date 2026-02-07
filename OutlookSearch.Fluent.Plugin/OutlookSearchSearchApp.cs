using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using Avalonia.Input;
using Blast.API.Search;
using Blast.API.Search.SearchOperations;
using Blast.Core;
using Blast.Core.Interfaces;
using Blast.Core.Objects;
using Blast.Core.Results;

namespace OutlookSearch.Fluent.Plugin;

public sealed class OutlookSearchSearchApp : ISearchApplication
{
    private const string SearchAppName = "Outlook Search";
    private const string MainTag = "outlook";
    private const string EmailTag = "email";
    private const string CalendarTag = "calendar";

    private readonly List<SearchTag> _defaultTags;
    private readonly List<SearchTag> _emailTags;
    private readonly List<SearchTag> _calendarTags;
    private readonly SearchApplicationInfo _applicationInfo;
    private readonly OutlookSearchSettingsPage _settingsPage;

    // Backends
    private OutlookInteropService? _comService;
    private GraphOutlookService? _graphService;
    private bool _comAvailable;
    private bool _comChecked;

    public OutlookSearchSearchApp()
    {
        var outlookTag = new SearchTag { Name = MainTag, Description = "search emails and events", IconGlyph = "\uE715" };
        var emailTag = new SearchTag { Name = EmailTag, Description = "search emails only", IconGlyph = "\uE8A8" };
        var calendarTag = new SearchTag { Name = CalendarTag, Description = "search calendar events only", IconGlyph = "\uE787" };

        _defaultTags = new List<SearchTag> { outlookTag, emailTag, calendarTag };
        _emailTags = new List<SearchTag> { emailTag, outlookTag };
        _calendarTags = new List<SearchTag> { calendarTag, outlookTag };

        _applicationInfo = new SearchApplicationInfo(
            SearchAppName,
            "Search Outlook emails and calendar events",
            new ObservableCollection<ISearchOperation>())
        {
            SearchTagOnly = true,
            MinimumSearchLength = 0,
            MinimumTagSearchLength = 0,
            SearchEmptyTextEmptyTag = false,
            IsProcessSearchEnabled = false,
            IsProcessSearchOffline = false,
            SearchAllTime = ApplicationSearchTime.Moderate,
            ApplicationIconGlyph = "\uE715",
            DefaultSearchTags = _defaultTags
        };

        _settingsPage = new OutlookSearchSettingsPage(_applicationInfo);
        _applicationInfo.SettingsPage = _settingsPage;
    }

    public async ValueTask LoadSearchApplicationAsync()
    {
        // Initialize Graph service early so persistent token cache is loaded.
        // Only attempt silent auth — never open browser on startup.
        try
        {
            _graphService = new GraphOutlookService();
            _graphService.Initialize(_settingsPage.GraphClientId);
            await _graphService.TrySilentAuthenticateAsync();
        }
        catch
        {
            // Non-fatal — user can sign in manually later
        }
    }

    public SearchApplicationInfo GetApplicationInfo() => _applicationInfo;

    public async IAsyncEnumerable<ISearchResult> SearchAsync(
        SearchRequest searchRequest,
        [EnumeratorCancellation] CancellationToken cancellationToken)
    {
        if (cancellationToken.IsCancellationRequested || searchRequest.SearchType == SearchType.SearchProcess)
            yield break;

        string tag = searchRequest.SearchedTag;
        bool isOutlookTag = string.Equals(tag, MainTag, StringComparison.OrdinalIgnoreCase);
        bool isEmailTag = string.Equals(tag, EmailTag, StringComparison.OrdinalIgnoreCase);
        bool isCalendarTag = string.Equals(tag, CalendarTag, StringComparison.OrdinalIgnoreCase);

        if (!isOutlookTag && !isEmailTag && !isCalendarTag)
            yield break;

        string searchedText = searchRequest.DisplayedSearchText?.Trim() ?? string.Empty;

        bool wantEmails = (isOutlookTag || isEmailTag) && _settingsPage.SearchEmails;
        bool wantEvents = (isOutlookTag || isCalendarTag) && _settingsPage.SearchEvents;
        bool wantActions = isOutlookTag && _settingsPage.ShowQuickActions;
        bool isEmptySearch = searchedText.Length == 0;

        // Quick actions — always available, self-handling via RunOperationFunc
        if (wantActions)
        {
            foreach (ISearchResult r in CreateQuickActionResults(searchedText))
            {
                if (cancellationToken.IsCancellationRequested) yield break;
                yield return r;
            }

            // Sign-out action when authenticated
            if (_graphService is { IsAuthenticated: true }
                && (searchedText.Length == 0
                    || "sign out".Contains(searchedText, StringComparison.OrdinalIgnoreCase)
                    || "logout".Contains(searchedText, StringComparison.OrdinalIgnoreCase)))
            {
                yield return CreateSignOutResult();
            }
        }

        // Determine backend
        EnsureBackend();
        bool hasCom = _comAvailable && _comService is { IsConnected: true };
        bool hasGraph = _graphService is { IsAuthenticated: true };

        // If Graph token expired, try silent refresh before falling back to sign-in
        if (!hasGraph && _graphService is { IsInitialized: true, AuthFailed: false })
        {
            try
            {
                hasGraph = await _graphService.TrySilentAuthenticateAsync(cancellationToken);
            }
            catch
            {
                // ignore
            }
        }

        // Prompt sign-in if no backend (self-handling operation, no auto-browser)
        if (!hasCom && !hasGraph)
        {
            if (!isEmptySearch && _graphService is not { AuthFailed: true })
                yield return CreateSignInResult();
            yield break;
        }

        // Upcoming events on empty search
        if (isEmptySearch && wantEvents && _settingsPage.ShowUpcomingEventsOnEmpty)
        {
            var upcoming = hasGraph
                ? await _graphService!.SearchEventsAsync("", _settingsPage.MaxEventResults, 0, _settingsPage.EventFutureDays, cancellationToken)
                : _comService!.SearchEvents("", _settingsPage.MaxEventResults, 0, _settingsPage.EventFutureDays);

            double s = 5.0;
            foreach (var evt in upcoming)
            {
                if (cancellationToken.IsCancellationRequested) yield break;
                yield return new OutlookEventSearchResult(evt, CreateEventOperations(evt), _calendarTags, s);
                s -= 0.1;
            }
            yield break;
        }

        if (isEmptySearch) yield break;

        // Search emails
        if (wantEmails)
        {
            var emails = hasGraph
                ? await _graphService!.SearchEmailsAsync(searchedText, _settingsPage.MaxEmailResults, _settingsPage.EmailSearchDaysBack, cancellationToken)
                : _comService!.SearchEmails(searchedText, _settingsPage.MaxEmailResults, _settingsPage.EmailSearchDaysBack);

            double s = 8.0;
            foreach (var email in emails)
            {
                if (cancellationToken.IsCancellationRequested) yield break;
                yield return new OutlookEmailSearchResult(email, CreateEmailOperations(email), _emailTags, s);
                s -= 0.1;
            }
        }

        // Search events
        if (wantEvents)
        {
            var events = hasGraph
                ? await _graphService!.SearchEventsAsync(searchedText, _settingsPage.MaxEventResults, _settingsPage.EventSearchDaysBack, _settingsPage.EventFutureDays, cancellationToken)
                : _comService!.SearchEvents(searchedText, _settingsPage.MaxEventResults, _settingsPage.EventSearchDaysBack, _settingsPage.EventFutureDays);

            double s = 6.0;
            foreach (var evt in events)
            {
                if (cancellationToken.IsCancellationRequested) yield break;
                yield return new OutlookEventSearchResult(evt, CreateEventOperations(evt), _calendarTags, s);
                s -= 0.1;
            }
        }
    }

    public ValueTask<IHandleResult> HandleSearchResult(ISearchResult searchResult) =>
        ValueTask.FromResult<IHandleResult>(new HandleResult(true, false));

    #region Operation factories — all use RunOperationFunc so they self-handle

    private ISearchResult CreateSignInResult()
    {
        var op = MakeOp("Sign In", "Sign in with your Microsoft account", "\uE77B", async _ =>
        {
            _graphService ??= new GraphOutlookService();
            if (!_graphService.IsInitialized)
                _graphService.Initialize(_settingsPage.GraphClientId);
            bool ok = await _graphService.TryAuthenticateAsync();
            return new HandleResult(ok, searchAgain: ok);
        });
        op.KeyGesture = new KeyGesture(Key.Enter);
        op.HideMainWindow = false;

        return new OutlookActionSearchResult(
            "Sign in to search Outlook",
            "Sign in with your Microsoft account to search emails and events",
            "\uE77B", "sign_in",
            new List<ISearchOperation> { op }, _defaultTags, 10.0);
    }

    private ISearchResult CreateSignOutResult()
    {
        var op = MakeOp("Sign Out", "Sign out of your Microsoft account", "\uF3B1", async _ =>
        {
            if (_graphService != null)
                await _graphService.SignOutAsync();
            return new HandleResult(true, searchAgain: true);
        });
        op.KeyGesture = new KeyGesture(Key.Enter);

        return new OutlookActionSearchResult(
            "Sign out of Outlook",
            "Sign out and clear cached credentials",
            "\uF3B1", "sign_out",
            new List<ISearchOperation> { op }, _defaultTags, 0.1);
    }

    private IEnumerable<ISearchResult> CreateQuickActionResults(string searchedText)
    {
        (string Name, string Desc, string Icon, string Id, string Uri)[] actions =
        [
            ("Compose New Email", "Open a new email compose window", "\uE70F", "compose", "mailto:"),
            ("Schedule New Meeting", "Schedule a new meeting in Outlook", "\uE787", "meeting", "https://outlook.office.com/calendar/0/deeplink/compose"),
            ("Open Inbox", "Open the Outlook inbox", "\uE715", "inbox", "ms-outlook:"),
            ("Open Calendar", "Open the Outlook calendar", "\uE787", "calendar", "https://outlook.office.com/calendar/view/month"),
        ];

        double score = 10.0;
        foreach (var (name, desc, icon, id, uri) in actions)
        {
            if (searchedText.Length > 0
                && !name.Contains(searchedText, StringComparison.OrdinalIgnoreCase)
                && !id.Contains(searchedText, StringComparison.OrdinalIgnoreCase))
                continue;

            string capturedUri = uri;
            var op = MakeOp(name, desc, icon, _ =>
            {
                bool ok = Launch(capturedUri);
                return ValueTask.FromResult<IHandleResult>(new HandleResult(ok, false, ok));
            });
            op.KeyGesture = new KeyGesture(Key.Enter);

            yield return new OutlookActionSearchResult(
                name, desc, icon, id,
                new List<ISearchOperation> { op }, _defaultTags, score);
            score -= 0.1;
        }
    }

    private List<ISearchOperation> CreateEmailOperations(OutlookEmailItem email)
    {
        var ops = new List<ISearchOperation>();

        if (_comAvailable && _comService is { IsConnected: true } && !string.IsNullOrEmpty(email.EntryId))
        {
            string eid = email.EntryId;
            var open = MakeOp("Open in Outlook", "Opens this email in Outlook", "\uE8A7", _ =>
            {
                bool ok = _comService!.OpenItem(eid);
                return ValueTask.FromResult<IHandleResult>(new HandleResult(ok, false, ok));
            });
            open.KeyGesture = new KeyGesture(Key.Enter);
            ops.Add(open);

            ops.Add(MakeOp("Reply", "Reply to this email", "\uE97A", _ =>
            {
                bool ok = _comService!.ReplyToEmail(eid);
                return ValueTask.FromResult<IHandleResult>(new HandleResult(ok, false, ok));
            }));
            ops.Add(MakeOp("Reply All", "Reply to all recipients", "\uE97B", _ =>
            {
                bool ok = _comService!.ReplyAllToEmail(eid);
                return ValueTask.FromResult<IHandleResult>(new HandleResult(ok, false, ok));
            }));
            ops.Add(MakeOp("Forward", "Forward this email", "\uE989", _ =>
            {
                bool ok = _comService!.ForwardEmail(eid);
                return ValueTask.FromResult<IHandleResult>(new HandleResult(ok, false, ok));
            }));
        }
        else if (!string.IsNullOrEmpty(email.WebLink))
        {
            string link = email.WebLink;
            var open = MakeOp("Open Email", "Opens this email", "\uE8A7", _ =>
            {
                Launch(link);
                return ValueTask.FromResult<IHandleResult>(new HandleResult(true, false, true));
            });
            open.KeyGesture = new KeyGesture(Key.Enter);
            ops.Add(open);
        }

        ops.Add(new CopySearchOperationSelfRun("Copy Subject"));
        return ops;
    }

    private List<ISearchOperation> CreateEventOperations(OutlookCalendarItem evt)
    {
        var ops = new List<ISearchOperation>();

        if (_comAvailable && _comService is { IsConnected: true } && !string.IsNullOrEmpty(evt.EntryId))
        {
            string eid = evt.EntryId;
            var open = MakeOp("Open in Outlook", "Opens this event in Outlook", "\uE8A7", _ =>
            {
                bool ok = _comService!.OpenItem(eid);
                return ValueTask.FromResult<IHandleResult>(new HandleResult(ok, false, ok));
            });
            open.KeyGesture = new KeyGesture(Key.Enter);
            ops.Add(open);
        }
        else if (!string.IsNullOrEmpty(evt.WebLink))
        {
            string link = evt.WebLink;
            var open = MakeOp("Open Event", "Opens this event", "\uE8A7", _ =>
            {
                Launch(link);
                return ValueTask.FromResult<IHandleResult>(new HandleResult(true, false, true));
            });
            open.KeyGesture = new KeyGesture(Key.Enter);
            ops.Add(open);
        }

        ops.Add(new CopySearchOperationSelfRun("Copy Title"));
        return ops;
    }

    #endregion

    #region Helpers

    private static ActionSearchOperation MakeOp(string name, string description, string iconGlyph,
        Func<ISearchResult, ValueTask<IHandleResult>> func)
    {
        return new ActionSearchOperation(func)
        {
            OperationName = name,
            Description = description,
            IconGlyph = iconGlyph,
            UniqueId = name + " " + description
        };
    }

    private static bool Launch(string uri)
    {
        try
        {
            Process.Start(new ProcessStartInfo(uri) { UseShellExecute = true });
            return true;
        }
        catch
        {
            return false;
        }
    }

    private void EnsureBackend()
    {
        if (!_comChecked)
        {
            _comChecked = true;
            Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
            _comAvailable = outlookType != null;
            if (_comAvailable)
            {
                _comService ??= new OutlookInteropService();
                _comService.TryConnect();
            }
        }
    }

    #endregion
}

