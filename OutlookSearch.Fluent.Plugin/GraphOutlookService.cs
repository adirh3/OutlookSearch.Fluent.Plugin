using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;

namespace OutlookSearch.Fluent.Plugin;

/// <summary>
/// Searches Outlook emails and calendar events via Microsoft Graph REST API.
/// Works with New Outlook (MSIX/Store) and any Microsoft 365 account.
/// Uses MSAL for OAuth2 authentication.
/// </summary>
internal sealed class GraphOutlookService
{
    // Well-known public client ID used by Microsoft Graph Command Line Tools.
    // Safe for public desktop apps; works with any org/personal Microsoft account.
    private const string DefaultClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
    private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";

    private static readonly string[] Scopes = ["Mail.Read", "Calendars.Read"];

    private readonly HttpClient _httpClient = new();
    private IPublicClientApplication? _msalApp;
    private string? _accessToken;
    private DateTimeOffset _tokenExpiry;
    private bool _authFailed;

    /// <summary>
    /// Whether the service has a valid access token.
    /// </summary>
    public bool IsAuthenticated => _accessToken != null && DateTimeOffset.UtcNow < _tokenExpiry;

    /// <summary>
    /// Whether Initialize() has been called.
    /// </summary>
    public bool IsInitialized => _msalApp != null;

    /// <summary>
    /// Whether authentication has been attempted and failed (don't retry automatically).
    /// </summary>
    public bool AuthFailed => _authFailed;

    /// <summary>
    /// Initialize MSAL app with persistent token cache. Call once.
    /// </summary>
    public void Initialize(string? clientId = null)
    {
        string effectiveClientId = string.IsNullOrWhiteSpace(clientId) ? DefaultClientId : clientId;

        _msalApp = PublicClientApplicationBuilder
            .Create(effectiveClientId)
            .WithAuthority("https://login.microsoftonline.com/common")
            .WithDefaultRedirectUri()
            .Build();

        // Register persistent file-based token cache so tokens survive app restarts
        RegisterTokenCacheAsync().GetAwaiter().GetResult();
    }

    private async Task RegisterTokenCacheAsync()
    {
        string cacheDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "FluentSearch", "OutlookPlugin");
        Directory.CreateDirectory(cacheDir);

        string cacheFile = "outlookplugin_msal.cache";
        var storageProperties = new StorageCreationPropertiesBuilder(cacheFile, cacheDir)
            .WithUnprotectedFile() // Fallback when DPAPI is unavailable (e.g., plugin host context)
            .Build();

        var cacheHelper = await MsalCacheHelper.CreateAsync(storageProperties);

        // Verify the cache helper can read/write; fall back to unprotected if DPAPI fails
        try
        {
            cacheHelper.VerifyPersistence();
        }
        catch
        {
            // DPAPI failed — rebuild with plaintext only
            storageProperties = new StorageCreationPropertiesBuilder(cacheFile, cacheDir)
                .WithUnprotectedFile()
                .Build();
            cacheHelper = await MsalCacheHelper.CreateAsync(storageProperties);
        }

        cacheHelper.RegisterCache(_msalApp!.UserTokenCache);
    }

    /// <summary>
    /// Attempt to acquire a token silently (from cache) only. Never opens browser.
    /// Returns true if a cached token was available.
    /// </summary>
    public async Task<bool> TrySilentAuthenticateAsync(CancellationToken cancellationToken = default)
    {
        if (IsAuthenticated)
            return true;

        if (_msalApp == null)
            return false;

        try
        {
            var accounts = await _msalApp.GetAccountsAsync();
            var account = accounts.FirstOrDefault();
            if (account != null)
            {
                var silentResult = await _msalApp.AcquireTokenSilent(Scopes, account)
                    .ExecuteAsync(cancellationToken);
                SetToken(silentResult);
                return true;
            }
        }
        catch
        {
            // Silent failed — user will need to sign in interactively
        }

        return false;
    }

    /// <summary>
    /// Attempt to acquire a token silently (from cache), falling back to interactive.
    /// Returns true if successful.
    /// </summary>
    public async Task<bool> TryAuthenticateAsync(CancellationToken cancellationToken = default)
    {
        if (IsAuthenticated)
            return true;

        if (_msalApp == null)
            Initialize();

        try
        {
            // Try silent first (cached token)
            var accounts = await _msalApp!.GetAccountsAsync();
            var account = accounts.FirstOrDefault();
            if (account != null)
            {
                var silentResult = await _msalApp.AcquireTokenSilent(Scopes, account)
                    .ExecuteAsync(cancellationToken);
                SetToken(silentResult);
                return true;
            }
        }
        catch (MsalUiRequiredException)
        {
            // Need interactive - fall through
        }
        catch
        {
            // Silent failed - fall through
        }

        try
        {
            // Interactive auth (opens browser)
            var interactiveResult = await _msalApp!.AcquireTokenInteractive(Scopes)
                .WithPrompt(Prompt.SelectAccount)
                .ExecuteAsync(cancellationToken);
            SetToken(interactiveResult);
            _authFailed = false;
            return true;
        }
        catch (MsalServiceException)
        {
            _authFailed = true;
            return false;
        }
        catch (OperationCanceledException)
        {
            return false;
        }
        catch
        {
            _authFailed = true;
            return false;
        }
    }

    /// <summary>
    /// Sign out: remove all cached accounts and clear token state.
    /// </summary>
    public async Task SignOutAsync()
    {
        if (_msalApp != null)
        {
            var accounts = await _msalApp.GetAccountsAsync();
            foreach (var account in accounts)
            {
                await _msalApp.RemoveAsync(account);
            }
        }

        _accessToken = null;
        _tokenExpiry = DateTimeOffset.MinValue;
        _authFailed = false;
    }

    /// <summary>
    /// Search emails via Microsoft Graph /me/messages.
    /// </summary>
    public async Task<List<OutlookEmailItem>> SearchEmailsAsync(
        string query, int maxResults, int searchDaysBack, CancellationToken cancellationToken)
    {
        var results = new List<OutlookEmailItem>();
        if (!IsAuthenticated || string.IsNullOrWhiteSpace(query))
            return results;

        try
        {
            // KQL field syntax: each field:term joined with OR
            // Graph $search requires the value wrapped in double quotes in the URL
            string escaped = query.Replace("'", "''").Replace("\"", "");
            string kql = $"subject:{escaped} OR from:{escaped} OR body:{escaped} OR to:{escaped}";
            string url =
                $"{GraphBaseUrl}/me/messages" +
                $"?$search=\"{Uri.EscapeDataString(kql)}\"" +
                $"&$top={maxResults}" +
                $"&$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments,isRead,importance,parentFolderId,webLink";

            var response = await GetAsync<GraphMessageListResponse>(url, cancellationToken);
            if (response?.Value == null)
                return results;

            foreach (var msg in response.Value)
            {
                if (cancellationToken.IsCancellationRequested)
                    break;

                results.Add(new OutlookEmailItem
                {
                    EntryId = msg.Id ?? "",
                    Subject = msg.Subject ?? "(No Subject)",
                    SenderName = msg.From?.EmailAddress?.Name ?? "",
                    SenderEmail = msg.From?.EmailAddress?.Address ?? "",
                    ReceivedTime = msg.ReceivedDateTime?.LocalDateTime ?? DateTime.MinValue,
                    BodyPreview = msg.BodyPreview ?? "",
                    HasAttachments = msg.HasAttachments,
                    IsRead = msg.IsRead,
                    ToRecipients = FormatRecipients(msg.ToRecipients),
                    CcRecipients = FormatRecipients(msg.CcRecipients),
                    Importance = msg.Importance?.ToLowerInvariant() switch
                    {
                        "high" => 2,
                        "low" => 0,
                        _ => 1
                    },
                    FolderName = "Inbox",
                    WebLink = msg.WebLink ?? ""
                });
            }
        }
        catch (OperationCanceledException)
        {
            // Expected
        }
        catch
        {
            // Return whatever we have
        }

        return results;
    }

    /// <summary>
    /// Search calendar events via Microsoft Graph /me/calendarView or /me/events.
    /// </summary>
    public async Task<List<OutlookCalendarItem>> SearchEventsAsync(
        string query, int maxResults, int searchDaysBack, int futureDays,
        CancellationToken cancellationToken)
    {
        var results = new List<OutlookCalendarItem>();
        if (!IsAuthenticated)
            return results;

        try
        {
            DateTime startDate = DateTime.UtcNow.AddDays(-searchDaysBack);
            DateTime endDate = DateTime.UtcNow.AddDays(futureDays);

            string url;
            // Graph calendarView doesn't support $filter=contains(), so fetch events and filter client-side
            url = $"{GraphBaseUrl}/me/calendarView" +
                  $"?startDateTime={startDate:yyyy-MM-ddTHH:mm:ssZ}" +
                  $"&endDateTime={endDate:yyyy-MM-ddTHH:mm:ssZ}" +
                  $"&$top={Math.Max(maxResults, 50)}" +
                  $"&$orderby=start/dateTime" +
                  $"&$select=id,subject,start,end,location,organizer,attendees,bodyPreview,isAllDay,webLink";

            var response = await GetAsync<GraphEventListResponse>(url, cancellationToken);
            if (response?.Value == null)
                return results;

            bool hasQuery = !string.IsNullOrWhiteSpace(query);
            foreach (var evt in response.Value)
            {
                if (cancellationToken.IsCancellationRequested)
                    break;

                // Client-side text filter
                if (hasQuery)
                {
                    bool matches = (evt.Subject ?? "").Contains(query, StringComparison.OrdinalIgnoreCase)
                                   || (evt.BodyPreview ?? "").Contains(query, StringComparison.OrdinalIgnoreCase)
                                   || (evt.Location?.DisplayName ?? "").Contains(query, StringComparison.OrdinalIgnoreCase);
                    if (!matches) continue;
                }

                if (results.Count >= maxResults)
                    break;

                results.Add(new OutlookCalendarItem
                {
                    EntryId = evt.Id ?? "",
                    Subject = evt.Subject ?? "",
                    StartTime = ParseGraphDateTime(evt.Start),
                    EndTime = ParseGraphDateTime(evt.End),
                    Location = evt.Location?.DisplayName ?? "",
                    Organizer = evt.Organizer?.EmailAddress?.Name ?? "",
                    RequiredAttendees = FormatAttendees(evt.Attendees, "required"),
                    OptionalAttendees = FormatAttendees(evt.Attendees, "optional"),
                    BodyPreview = evt.BodyPreview ?? "",
                    IsAllDayEvent = evt.IsAllDay,
                    WebLink = evt.WebLink ?? ""
                });
            }
        }
        catch (OperationCanceledException)
        {
            // Expected
        }
        catch
        {
            // Return whatever we have
        }

        return results;
    }

    private void SetToken(AuthenticationResult result)
    {
        _accessToken = result.AccessToken;
        _tokenExpiry = result.ExpiresOn;
    }

    private async Task<T?> GetAsync<T>(string url, CancellationToken ct)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);
        request.Headers.Add("Prefer", "outlook.body-type=\"text\"");

        using var response = await _httpClient.SendAsync(request, ct);
        if (!response.IsSuccessStatusCode)
            return default;

        return await response.Content.ReadFromJsonAsync<T>(GraphJsonOptions, ct);
    }

    private static string FormatRecipients(List<GraphRecipient>? recipients)
    {
        if (recipients == null || recipients.Count == 0)
            return string.Empty;
        return string.Join("; ", recipients
            .Where(r => r.EmailAddress != null)
            .Select(r => string.IsNullOrWhiteSpace(r.EmailAddress!.Name)
                ? r.EmailAddress.Address ?? ""
                : r.EmailAddress.Name));
    }

    private static string FormatAttendees(List<GraphAttendee>? attendees, string type)
    {
        if (attendees == null || attendees.Count == 0)
            return string.Empty;
        return string.Join("; ", attendees
            .Where(a => string.Equals(a.Type, type, StringComparison.OrdinalIgnoreCase)
                        && a.EmailAddress != null)
            .Select(a => string.IsNullOrWhiteSpace(a.EmailAddress!.Name)
                ? a.EmailAddress.Address ?? ""
                : a.EmailAddress.Name));
    }

    private static DateTime ParseGraphDateTime(GraphDateTimeTimeZone? dt)
    {
        if (dt?.DateTime == null)
            return DateTime.MinValue;

        if (DateTime.TryParse(dt.DateTime, out DateTime parsed))
            return parsed;
        return DateTime.MinValue;
    }

    private static readonly JsonSerializerOptions GraphJsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    #region Graph API DTO types

    internal sealed class GraphMessageListResponse
    {
        [JsonPropertyName("value")]
        public List<GraphMessage>? Value { get; set; }
    }

    internal sealed class GraphMessage
    {
        public string? Id { get; set; }
        public string? Subject { get; set; }
        public GraphRecipient? From { get; set; }
        public List<GraphRecipient>? ToRecipients { get; set; }
        public List<GraphRecipient>? CcRecipients { get; set; }
        public DateTimeOffset? ReceivedDateTime { get; set; }
        public string? BodyPreview { get; set; }
        public bool HasAttachments { get; set; }
        public bool IsRead { get; set; }
        public string? Importance { get; set; }
        public string? WebLink { get; set; }
    }

    internal sealed class GraphEventListResponse
    {
        [JsonPropertyName("value")]
        public List<GraphEvent>? Value { get; set; }
    }

    internal sealed class GraphEvent
    {
        public string? Id { get; set; }
        public string? Subject { get; set; }
        public GraphDateTimeTimeZone? Start { get; set; }
        public GraphDateTimeTimeZone? End { get; set; }
        public GraphLocation? Location { get; set; }
        public GraphRecipient? Organizer { get; set; }
        public List<GraphAttendee>? Attendees { get; set; }
        public string? BodyPreview { get; set; }
        public bool IsAllDay { get; set; }
        public string? WebLink { get; set; }
    }

    internal sealed class GraphRecipient
    {
        public GraphEmailAddress? EmailAddress { get; set; }
    }

    internal sealed class GraphAttendee
    {
        public GraphEmailAddress? EmailAddress { get; set; }
        public string? Type { get; set; }
    }

    internal sealed class GraphEmailAddress
    {
        public string? Name { get; set; }
        public string? Address { get; set; }
    }

    internal sealed class GraphDateTimeTimeZone
    {
        public string? DateTime { get; set; }
        public string? TimeZone { get; set; }
    }

    internal sealed class GraphLocation
    {
        public string? DisplayName { get; set; }
    }

    #endregion
}
