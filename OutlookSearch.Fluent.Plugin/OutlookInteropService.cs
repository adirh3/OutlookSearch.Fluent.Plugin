using System.Runtime.InteropServices;

namespace OutlookSearch.Fluent.Plugin;

/// <summary>
/// Provides 100% local access to Outlook data via COM late-binding (dynamic dispatch).
/// No compile-time dependency on office type libraries – works with any Outlook desktop version.
/// </summary>
internal sealed class OutlookInteropService : IDisposable
{
    // Outlook OlItemType constants
    private const int OlMail = 0;
    private const int OlAppointment = 1;

    // Outlook OlDefaultFolders constants
    private const int OlFolderInbox = 6;
    private const int OlFolderCalendar = 9;
    private const int OlFolderSentMail = 5;

    // Outlook OlMeetingStatus
    private const int OlMeeting = 1;

    private dynamic? _outlookApp;
    private dynamic? _namespace;
    private bool _disposed;

    /// <summary>
    /// Whether the connection to Outlook is currently active.
    /// </summary>
    public bool IsConnected => _outlookApp != null;

    /// <summary>
    /// Try to attach to a running Outlook instance or start a new one.
    /// Returns true if successful.
    /// </summary>
    public bool TryConnect()
    {
        if (_outlookApp != null)
            return true;

        try
        {
            // Try to attach to a running instance first via ROT (Running Object Table)
            _outlookApp = GetActiveOutlookInstance();
        }
        catch
        {
            _outlookApp = null;
        }

        if (_outlookApp == null)
        {
            try
            {
                // Start a new instance
                Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null)
                    return false;
                _outlookApp = Activator.CreateInstance(outlookType);
            }
            catch
            {
                _outlookApp = null;
                return false;
            }
        }

        if (_outlookApp == null)
            return false;

        try
        {
            _namespace = _outlookApp.GetNamespace("MAPI");
            return true;
        }
        catch
        {
            _outlookApp = null;
            return false;
        }
    }

    /// <summary>
    /// Gets a running Outlook.Application instance.
    /// Replacement for Marshal.GetActiveObject which is not available in .NET 5+.
    /// </summary>
    private static object? GetActiveOutlookInstance()
    {
        try
        {
            Type? type = Type.GetTypeFromProgID("Outlook.Application");
            if (type == null)
                return null;
            int result = CLSIDFromProgID("Outlook.Application", out Guid clsid);
            if (result != 0)
                return null;
            result = OleGetActiveObject(ref clsid, IntPtr.Zero, out object activeObj);
            return result == 0 ? activeObj : null;
        }
        catch
        {
            return null;
        }
    }

    [DllImport("ole32.dll", EntryPoint = "CLSIDFromProgID")]
    private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

    [DllImport("oleaut32.dll", EntryPoint = "GetActiveObject")]
    private static extern int OleGetActiveObject(ref Guid clsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object obj);

    /// <summary>
    /// Searches emails in Inbox and Sent Mail using Outlook DASL filter.
    /// Returns up to <paramref name="maxResults"/> items.
    /// </summary>
    public List<OutlookEmailItem> SearchEmails(string query, int maxResults, int searchDaysBack)
    {
        var results = new List<OutlookEmailItem>();
        if (_namespace == null || string.IsNullOrWhiteSpace(query))
            return results;

        try
        {
            int[] folderIds = [OlFolderInbox, OlFolderSentMail];
            foreach (int folderId in folderIds)
            {
                if (results.Count >= maxResults)
                    break;

                try
                {
                    SearchEmailsInFolder(folderId, query, maxResults - results.Count, searchDaysBack, results);
                }
                catch
                {
                    // Skip folder if it can't be accessed
                }
            }
        }
        catch
        {
            // Return whatever we have
        }

        return results;
    }

    /// <summary>
    /// Searches calendar events using Outlook Restrict filter.
    /// </summary>
    public List<OutlookCalendarItem> SearchEvents(string query, int maxResults, int searchDaysBack, int futureDays)
    {
        var results = new List<OutlookCalendarItem>();
        if (_namespace == null)
            return results;

        try
        {
            dynamic folder = _namespace.GetDefaultFolder(OlFolderCalendar);
            dynamic items = folder.Items;
            items.Sort("[Start]", true); // Most recent first
            items.IncludeRecurrences = true;

            // Date range filter
            DateTime startDate = DateTime.Now.AddDays(-searchDaysBack);
            DateTime endDate = DateTime.Now.AddDays(futureDays);
            string restrict =
                $"[Start] >= '{startDate:g}' AND [Start] <= '{endDate:g}'";

            dynamic filteredItems = items.Restrict(restrict);

            foreach (dynamic item in filteredItems)
            {
                if (results.Count >= maxResults)
                    break;

                try
                {
                    string subject = (string)(item.Subject ?? "");
                    string location = "";
                    try { location = item.Location ?? ""; } catch { /* ignored */ }
                    string body = "";
                    try { body = item.Body ?? ""; } catch { /* ignored */ }

                    bool matches = string.IsNullOrEmpty(query)
                                   || subject.Contains(query, StringComparison.OrdinalIgnoreCase)
                                   || location.Contains(query, StringComparison.OrdinalIgnoreCase)
                                   || body.Contains(query, StringComparison.OrdinalIgnoreCase);

                    if (!matches)
                        continue;

                    string organizer = "";
                    try { organizer = item.Organizer ?? ""; } catch { /* ignored */ }

                    string requiredAttendees = "";
                    try { requiredAttendees = item.RequiredAttendees ?? ""; } catch { /* ignored */ }

                    string optionalAttendees = "";
                    try { optionalAttendees = item.OptionalAttendees ?? ""; } catch { /* ignored */ }

                    bool isAllDayEvent = false;
                    try { isAllDayEvent = item.AllDayEvent; } catch { /* ignored */ }

                    string entryId = item.EntryID;

                    results.Add(new OutlookCalendarItem
                    {
                        EntryId = entryId,
                        Subject = subject,
                        StartTime = (DateTime)item.Start,
                        EndTime = (DateTime)item.End,
                        Location = location,
                        Organizer = organizer,
                        RequiredAttendees = requiredAttendees,
                        OptionalAttendees = optionalAttendees,
                        BodyPreview = TruncateBody(body, 500),
                        IsAllDayEvent = isAllDayEvent
                    });
                }
                catch
                {
                    // Skip problematic items
                }
                finally
                {
                    ReleaseCom(item);
                }
            }

            ReleaseCom(filteredItems);
            ReleaseCom(items);
            ReleaseCom(folder);
        }
        catch
        {
            // Return whatever we have
        }

        return results;
    }

    /// <summary>
    /// Opens a mail item by EntryID in Outlook.
    /// </summary>
    public bool OpenItem(string entryId)
    {
        if (_namespace == null || string.IsNullOrEmpty(entryId))
            return false;
        try
        {
            dynamic item = _namespace!.GetItemFromID(entryId);
            item.Display();
            ReleaseCom(item);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Creates a reply to the email with the given EntryID.
    /// </summary>
    public bool ReplyToEmail(string entryId)
    {
        if (_namespace == null || string.IsNullOrEmpty(entryId))
            return false;
        try
        {
            dynamic item = _namespace!.GetItemFromID(entryId);
            dynamic reply = item.Reply();
            reply.Display();
            ReleaseCom(reply);
            ReleaseCom(item);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Creates a reply-all to the email with the given EntryID.
    /// </summary>
    public bool ReplyAllToEmail(string entryId)
    {
        if (_namespace == null || string.IsNullOrEmpty(entryId))
            return false;
        try
        {
            dynamic item = _namespace!.GetItemFromID(entryId);
            dynamic reply = item.ReplyAll();
            reply.Display();
            ReleaseCom(reply);
            ReleaseCom(item);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Forwards the email with the given EntryID.
    /// </summary>
    public bool ForwardEmail(string entryId)
    {
        if (_namespace == null || string.IsNullOrEmpty(entryId))
            return false;
        try
        {
            dynamic item = _namespace!.GetItemFromID(entryId);
            dynamic fwd = item.Forward();
            fwd.Display();
            ReleaseCom(fwd);
            ReleaseCom(item);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Opens a new blank email compose window.
    /// </summary>
    public bool ComposeNewEmail()
    {
        if (_outlookApp == null)
            return false;
        try
        {
            dynamic mail = _outlookApp.CreateItem(OlMail);
            mail.Display();
            ReleaseCom(mail);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Opens a new meeting scheduling window.
    /// </summary>
    public bool ScheduleNewMeeting()
    {
        if (_outlookApp == null)
            return false;
        try
        {
            dynamic appointment = _outlookApp.CreateItem(OlAppointment);
            appointment.MeetingStatus = OlMeeting;
            appointment.Display();
            ReleaseCom(appointment);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Opens the Outlook calendar view.
    /// </summary>
    public bool OpenCalendar()
    {
        if (_namespace == null)
            return false;
        try
        {
            dynamic folder = _namespace.GetDefaultFolder(OlFolderCalendar);
            folder.Display();
            ReleaseCom(folder);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Opens the Outlook inbox view.
    /// </summary>
    public bool OpenInbox()
    {
        if (_namespace == null)
            return false;
        try
        {
            dynamic folder = _namespace.GetDefaultFolder(OlFolderInbox);
            folder.Display();
            ReleaseCom(folder);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private void SearchEmailsInFolder(int folderId, string query, int maxResults, int searchDaysBack,
        List<OutlookEmailItem> results)
    {
        dynamic folder = _namespace!.GetDefaultFolder(folderId);
        dynamic items = folder.Items;
        items.Sort("[ReceivedTime]", true); // Most recent first

        // Build DASL filter: search Subject and SenderName containing the query
        // Also filter by date range for performance
        DateTime cutoffDate = DateTime.Now.AddDays(-searchDaysBack);
        string escapedQuery = query.Replace("'", "''");

        string filter =
            $"@SQL=(" +
            $"\"urn:schemas:httpmail:subject\" LIKE '%{escapedQuery}%'" +
            $" OR \"urn:schemas:httpmail:fromname\" LIKE '%{escapedQuery}%'" +
            $" OR \"urn:schemas:httpmail:fromemail\" LIKE '%{escapedQuery}%'" +
            $") AND \"urn:schemas:httpmail:datereceived\" >= '{cutoffDate:yyyy-MM-ddTHH:mm:ssZ}'";

        dynamic filteredItems = items.Restrict(filter);

        foreach (dynamic item in filteredItems)
        {
            if (results.Count >= maxResults)
                break;

            try
            {
                string subject = (string)(item.Subject ?? "(No Subject)");
                string senderName = "";
                string senderEmail = "";
                try { senderName = item.SenderName ?? ""; } catch { /* ignored */ }  
                try { senderEmail = item.SenderEmailAddress ?? ""; } catch { /* ignored */ }
                DateTime receivedTime = DateTime.MinValue;
                try { receivedTime = item.ReceivedTime; } catch { /* ignored */ }

                string bodyPreview = "";
                try { bodyPreview = TruncateBody(item.Body ?? "", 300); } catch { /* ignored */ }

                string htmlBody = "";
                try { htmlBody = item.HTMLBody ?? ""; } catch { /* ignored */ }

                bool hasAttachments = false;
                try { hasAttachments = item.Attachments?.Count > 0; } catch { /* ignored */ }

                bool isRead = true;
                try { isRead = !item.UnRead; } catch { /* ignored */ }

                string entryId = item.EntryID;
                string toRecipients = "";
                try { toRecipients = item.To ?? ""; } catch { /* ignored */ }
                string ccRecipients = "";
                try { ccRecipients = item.CC ?? ""; } catch { /* ignored */ }

                int importance = 1; // normal
                try { importance = (int)item.Importance; } catch { /* ignored */ }

                string folderName = folderId == OlFolderSentMail ? "Sent" : "Inbox";

                results.Add(new OutlookEmailItem
                {
                    EntryId = entryId,
                    Subject = subject,
                    SenderName = senderName,
                    SenderEmail = senderEmail,
                    ReceivedTime = receivedTime,
                    BodyPreview = bodyPreview,
                    HtmlBody = htmlBody,
                    HasAttachments = hasAttachments,
                    IsRead = isRead,
                    ToRecipients = toRecipients,
                    CcRecipients = ccRecipients,
                    Importance = importance,
                    FolderName = folderName
                });
            }
            catch
            {
                // Skip problematic items
            }
            finally
            {
                ReleaseCom(item);
            }
        }

        ReleaseCom(filteredItems);
        ReleaseCom(items);
        ReleaseCom(folder);
    }

    private static string TruncateBody(string body, int maxLength)
    {
        if (string.IsNullOrEmpty(body))
            return string.Empty;

        // Clean up excessive whitespace
        string cleaned = System.Text.RegularExpressions.Regex.Replace(body, @"\s+", " ").Trim();
        return cleaned.Length <= maxLength
            ? cleaned
            : cleaned[..maxLength] + "...";
    }

    private static void ReleaseCom(object? obj)
    {
        if (obj != null)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
            }
            catch
            {
                // ignored
            }
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        _namespace = null;
        if (_outlookApp != null)
        {
            // Don't quit Outlook—just release our reference
            ReleaseCom(_outlookApp);
            _outlookApp = null;
        }
    }
}

/// <summary>
/// Represents a local Outlook email item.
/// </summary>
internal sealed class OutlookEmailItem
{
    public string EntryId { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public string SenderName { get; set; } = string.Empty;
    public string SenderEmail { get; set; } = string.Empty;
    public DateTime ReceivedTime { get; set; }
    public string BodyPreview { get; set; } = string.Empty;
    public string HtmlBody { get; set; } = string.Empty;
    public bool HasAttachments { get; set; }
    public bool IsRead { get; set; }
    public string ToRecipients { get; set; } = string.Empty;
    public string CcRecipients { get; set; } = string.Empty;
    public int Importance { get; set; } = 1;
    public string FolderName { get; set; } = string.Empty;
    public string WebLink { get; set; } = string.Empty;
}

/// <summary>
/// Represents a local Outlook calendar item.
/// </summary>
internal sealed class OutlookCalendarItem
{
    public string EntryId { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public DateTime StartTime { get; set; }
    public DateTime EndTime { get; set; }
    public string Location { get; set; } = string.Empty;
    public string Organizer { get; set; } = string.Empty;
    public string RequiredAttendees { get; set; } = string.Empty;
    public string OptionalAttendees { get; set; } = string.Empty;
    public string BodyPreview { get; set; } = string.Empty;
    public bool IsAllDayEvent { get; set; }
    public string WebLink { get; set; } = string.Empty;
}
