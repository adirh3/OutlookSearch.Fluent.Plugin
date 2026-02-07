using System.Collections.ObjectModel;
using Blast.Core.Interfaces;
using Blast.Core.Objects;
using Blast.Core.Results;

namespace OutlookSearch.Fluent.Plugin;

/// <summary>
/// Search result representing an Outlook email.
/// </summary>
public sealed class OutlookEmailSearchResult : SearchResultBase
{
    internal string EntryId { get; }
    internal string WebLink { get; }
    internal string SenderName { get; }
    internal string SenderEmail { get; }
    internal DateTime ReceivedTime { get; }
    internal string BodyPreview { get; }
    internal string HtmlBody { get; }
    internal bool HasAttachments { get; }
    internal bool IsRead { get; }
    internal string ToRecipients { get; }
    internal string CcRecipients { get; }
    internal int Importance { get; }
    internal string FolderName { get; }

    internal OutlookEmailSearchResult(
        OutlookEmailItem email,
        IList<ISearchOperation> supportedOperations,
        ICollection<SearchTag> tags,
        double score) : base(
        searchApp: "Outlook Search",
        resultName: email.Subject,
        searchedText: null,
        resultType: "Email",
        score: score,
        supportedOperations: supportedOperations,
        tags: tags)
    {
        EntryId = email.EntryId;
        WebLink = email.WebLink;
        SenderName = email.SenderName;
        SenderEmail = email.SenderEmail;
        ReceivedTime = email.ReceivedTime;
        BodyPreview = email.BodyPreview;
        HtmlBody = email.HtmlBody;
        HasAttachments = email.HasAttachments;
        IsRead = email.IsRead;
        ToRecipients = email.ToRecipients;
        CcRecipients = email.CcRecipients;
        Importance = email.Importance;
        FolderName = email.FolderName;

        // Display
        DisplayedName = email.Subject;
        string timeStr = FormatReceivedTime(email.ReceivedTime);
        string attachmentStr = email.HasAttachments ? " \U0001F4CE" : "";
        string preview = string.IsNullOrWhiteSpace(email.BodyPreview)
            ? ""
            : " \u2014 " + email.BodyPreview.ReplaceLineEndings(" ").Trim();
        string info = $"{email.SenderName} \u00B7 {timeStr}{attachmentStr}{preview}";
        // Truncate to reasonable length for the UI
        AdditionalInformation = info.Length > 120 ? info[..117] + "..." : info;
        ShowAdditionalInformation = true;

        // Icon: unread = bold filled envelope, read = open envelope
        IconGlyph = email.IsRead ? "\uE8C3" : "\uE8A8";
        UseIconGlyph = true;

        // Unique IDs
        SearchObjectId = email.EntryId;
        PinUniqueId = "OutlookEmail_" + email.EntryId;

        // Information elements for preview panel
        var infoElements = new List<InformationElement>
        {
            new("From", $"{email.SenderName} <{email.SenderEmail}>"),
            new("To", email.ToRecipients),
            new("Date", email.ReceivedTime.ToString("f")),
            new("Folder", email.FolderName)
        };

        if (!string.IsNullOrWhiteSpace(email.CcRecipients))
            infoElements.Add(new InformationElement("CC", email.CcRecipients));

        if (email.HasAttachments)
            infoElements.Add(new InformationElement("Attachments", "Yes"));

        string importanceStr = email.Importance switch
        {
            0 => "Low",
            2 => "High",
            _ => "Normal"
        };
        if (email.Importance != 1) // Only show if not Normal
            infoElements.Add(new InformationElement("Importance", importanceStr));

        if (!string.IsNullOrWhiteSpace(email.BodyPreview))
            infoElements.Add(new InformationElement("Preview", email.BodyPreview));

        InformationElements = infoElements;

        // Set group name
        GroupName = "Emails";

        // Allow caching for faster subsequent searches
        ShouldCacheResult = true;
    }

    protected override void OnSelectedSearchResultChanged()
    {
    }

    private static string FormatReceivedTime(DateTime receivedTime)
    {
        if (receivedTime == DateTime.MinValue)
            return "Unknown";

        TimeSpan diff = DateTime.Now - receivedTime;

        if (diff.TotalMinutes < 1)
            return "Just now";
        if (diff.TotalMinutes < 60)
            return $"{(int)diff.TotalMinutes}m ago";
        if (diff.TotalHours < 24)
            return $"{(int)diff.TotalHours}h ago";
        if (diff.TotalDays < 7)
            return $"{(int)diff.TotalDays}d ago";
        if (receivedTime.Year == DateTime.Now.Year)
            return receivedTime.ToString("MMM d");
        return receivedTime.ToString("MMM d, yyyy");
    }
}

/// <summary>
/// Search result representing an Outlook calendar event.
/// </summary>
public sealed class OutlookEventSearchResult : SearchResultBase
{
    internal string EntryId { get; }
    internal string WebLink { get; }
    internal DateTime StartTime { get; }
    internal DateTime EndTime { get; }
    internal string Location { get; }
    internal string Organizer { get; }
    internal string RequiredAttendees { get; }
    internal string OptionalAttendees { get; }
    internal string EventBodyPreview { get; }
    internal bool IsAllDayEvent { get; }

    internal OutlookEventSearchResult(
        OutlookCalendarItem calendarItem,
        IList<ISearchOperation> supportedOperations,
        ICollection<SearchTag> tags,
        double score) : base(
        searchApp: "Outlook Search",
        resultName: calendarItem.Subject,
        searchedText: null,
        resultType: "Event",
        score: score,
        supportedOperations: supportedOperations,
        tags: tags)
    {
        EntryId = calendarItem.EntryId;
        WebLink = calendarItem.WebLink;
        StartTime = calendarItem.StartTime;
        EndTime = calendarItem.EndTime;
        Location = calendarItem.Location;
        Organizer = calendarItem.Organizer;
        RequiredAttendees = calendarItem.RequiredAttendees;
        OptionalAttendees = calendarItem.OptionalAttendees;
        EventBodyPreview = calendarItem.BodyPreview;
        IsAllDayEvent = calendarItem.IsAllDayEvent;

        // Display
        DisplayedName = calendarItem.Subject;
        string timeRange = FormatEventTime(calendarItem);
        string locationStr = string.IsNullOrWhiteSpace(calendarItem.Location)
            ? ""
            : $" · {calendarItem.Location}";
        AdditionalInformation = $"{timeRange}{locationStr}";
        ShowAdditionalInformation = true;

        // Calendar icon
        IconGlyph = "\uE787"; // Calendar
        UseIconGlyph = true;

        // Unique IDs
        SearchObjectId = calendarItem.EntryId;
        PinUniqueId = "OutlookEvent_" + calendarItem.EntryId;

        // Information elements
        var infoElements = new List<InformationElement>
        {
            new("When", timeRange),
        };

        if (!string.IsNullOrWhiteSpace(calendarItem.Location))
            infoElements.Add(new InformationElement("Location", calendarItem.Location));

        if (!string.IsNullOrWhiteSpace(calendarItem.Organizer))
            infoElements.Add(new InformationElement("Organizer", calendarItem.Organizer));

        if (!string.IsNullOrWhiteSpace(calendarItem.RequiredAttendees))
            infoElements.Add(new InformationElement("Attendees", calendarItem.RequiredAttendees));

        if (!string.IsNullOrWhiteSpace(calendarItem.OptionalAttendees))
            infoElements.Add(new InformationElement("Optional", calendarItem.OptionalAttendees));

        if (!string.IsNullOrWhiteSpace(calendarItem.BodyPreview))
            infoElements.Add(new InformationElement("Details", calendarItem.BodyPreview));

        // Add countdown/status info
        string status = GetEventStatus(calendarItem);
        if (!string.IsNullOrEmpty(status))
            infoElements.Insert(0, new InformationElement("Status", status));

        InformationElements = infoElements;

        // Set group name
        GroupName = "Events";

        ShouldCacheResult = true;
    }

    protected override void OnSelectedSearchResultChanged()
    {
    }

    private static string FormatEventTime(OutlookCalendarItem item)
    {
        if (item.IsAllDayEvent)
        {
            if (item.StartTime.Date == item.EndTime.Date || item.EndTime == item.StartTime.AddDays(1))
                return $"{item.StartTime:ddd, MMM d} (All Day)";
            return $"{item.StartTime:ddd, MMM d} - {item.EndTime.AddDays(-1):ddd, MMM d} (All Day)";
        }

        if (item.StartTime.Date == item.EndTime.Date)
            return $"{item.StartTime:ddd, MMM d} {item.StartTime:h:mm tt} - {item.EndTime:h:mm tt}";

        return $"{item.StartTime:ddd, MMM d h:mm tt} - {item.EndTime:ddd, MMM d h:mm tt}";
    }

    private static string GetEventStatus(OutlookCalendarItem item)
    {
        DateTime now = DateTime.Now;
        if (now >= item.StartTime && now <= item.EndTime)
            return "\u25CF Happening now";
        if (item.StartTime > now)
        {
            TimeSpan until = item.StartTime - now;
            if (until.TotalMinutes < 60)
                return $"In {(int)until.TotalMinutes} minutes";
            if (until.TotalHours < 24)
                return $"In {(int)until.TotalHours} hours";
            if (until.TotalDays < 7)
                return $"In {(int)until.TotalDays} days";
        }

        return string.Empty;
    }
}

/// <summary>
/// Search result representing a quick action (e.g., New Email, New Meeting).
/// </summary>
public sealed class OutlookActionSearchResult : SearchResultBase
{
    internal string ActionId { get; }

    public OutlookActionSearchResult(
        string actionName,
        string description,
        string iconGlyph,
        string actionId,
        IList<ISearchOperation> supportedOperations,
        ICollection<SearchTag> tags,
        double score) : base(
        searchApp: "Outlook Search",
        resultName: actionName,
        searchedText: null,
        resultType: "Action",
        score: score,
        supportedOperations: supportedOperations,
        tags: tags)
    {
        ActionId = actionId;
        DisplayedName = actionName;
        AdditionalInformation = description;
        ShowAdditionalInformation = true;
        IconGlyph = iconGlyph;
        UseIconGlyph = true;
        SearchObjectId = "OutlookAction_" + actionId;
        PinUniqueId = "OutlookAction_" + actionId;
        GroupName = "Quick Actions";
        DisableMachineLearning = true;

        InformationElements = new List<InformationElement>
        {
            new("Action", description)
        };
    }

    protected override void OnSelectedSearchResultChanged()
    {
    }
}
