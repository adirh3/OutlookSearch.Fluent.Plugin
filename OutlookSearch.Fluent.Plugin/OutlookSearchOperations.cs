using Avalonia.Input;
using Blast.Core.Interfaces;
using Blast.Core.Results;

namespace OutlookSearch.Fluent.Plugin;

/// <summary>
/// Opens the Outlook item (email/event) in the Outlook desktop app.
/// </summary>
internal sealed class OpenInOutlookOperation : SearchOperationBase
{
    public static readonly OpenInOutlookOperation Instance = new();

    public OpenInOutlookOperation() : base("Open in Outlook", "Opens this item in Outlook", "\uE8A7")
    {
        KeyGesture = new KeyGesture(Key.Enter);
    }
}

/// <summary>
/// Replies to an email.
/// </summary>
internal sealed class ReplyOperation : SearchOperationBase
{
    public static readonly ReplyOperation Instance = new();

    public ReplyOperation() : base("Reply", "Reply to this email", "\uE97A")
    {
        KeyGesture = new KeyGesture(Key.R, KeyModifiers.Control);
    }
}

/// <summary>
/// Replies to all recipients of an email.
/// </summary>
internal sealed class ReplyAllOperation : SearchOperationBase
{
    public static readonly ReplyAllOperation Instance = new();

    public ReplyAllOperation() : base("Reply All", "Reply to all recipients", "\uE97B")
    {
        KeyGesture = new KeyGesture(Key.R, KeyModifiers.Control | KeyModifiers.Shift);
    }
}

/// <summary>
/// Forwards an email.
/// </summary>
internal sealed class ForwardOperation : SearchOperationBase
{
    public static readonly ForwardOperation Instance = new();

    public ForwardOperation() : base("Forward", "Forward this email", "\uE989")
    {
        KeyGesture = new KeyGesture(Key.F, KeyModifiers.Control);
    }
}

/// <summary>
/// Opens a new email compose window (used for quick actions).
/// </summary>
internal sealed class ComposeNewEmailOperation : SearchOperationBase
{
    public static readonly ComposeNewEmailOperation Instance = new();

    public ComposeNewEmailOperation() : base("Compose New Email", "Open a new email compose window", "\uE70F")
    {
        KeyGesture = new KeyGesture(Key.N, KeyModifiers.Control);
    }
}

/// <summary>
/// Opens a new meeting scheduling window (used for quick actions).
/// </summary>
internal sealed class ScheduleMeetingOperation : SearchOperationBase
{
    public static readonly ScheduleMeetingOperation Instance = new();

    public ScheduleMeetingOperation() : base("Schedule Meeting", "Open a new meeting window", "\uE787")
    {
        KeyGesture = new KeyGesture(Key.N, KeyModifiers.Control | KeyModifiers.Shift);
    }
}

/// <summary>
/// Opens Outlook Calendar view.
/// </summary>
internal sealed class OpenCalendarOperation : SearchOperationBase
{
    public static readonly OpenCalendarOperation Instance = new();

    public OpenCalendarOperation() : base("Open Calendar", "Open Outlook Calendar", "\uE787")
    {
    }
}

/// <summary>
/// Opens Outlook Inbox view.
/// </summary>
internal sealed class OpenInboxOperation : SearchOperationBase
{
    public static readonly OpenInboxOperation Instance = new();

    public OpenInboxOperation() : base("Open Inbox", "Open Outlook Inbox", "\uE715")
    {
    }
}
