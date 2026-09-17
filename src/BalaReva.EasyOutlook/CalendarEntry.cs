namespace BalaReva.EasyOutlook;

/// <summary>The details of an appointment to add to the calendar.</summary>
/// <remarks>
/// Not part of the published surface. It exists so <see cref="IOutlookService"/> takes
/// one parameter instead of a dozen, which is what lets the activities be tested
/// without Outlook.
/// </remarks>
public class CalendarEntry
{
    /// <summary>Outlook account whose calendar to add to. Empty means the default.</summary>
    public string Account { get; set; } = string.Empty;

    /// <summary>Subject line.</summary>
    public string Subject { get; set; } = string.Empty;

    /// <summary>Body text.</summary>
    public string Body { get; set; } = string.Empty;

    /// <summary>Where it takes place.</summary>
    public string Location { get; set; } = string.Empty;

    /// <summary>When it starts.</summary>
    public DateTime Start { get; set; }

    /// <summary>When it ends.</summary>
    public DateTime End { get; set; }

    /// <summary>True for an all-day entry, in which case the times are ignored.</summary>
    public bool AllDayEvent { get; set; }

    /// <summary>Minutes before the start to show a reminder.</summary>
    public int ReminderMinutes { get; set; }

    /// <summary>How it shows on the calendar.</summary>
    public BusyStatusEnum BusyStatus { get; set; } = BusyStatusEnum.Busy;

    /// <summary>Importance flag.</summary>
    public ImportanceEnum Importance { get; set; } = ImportanceEnum.Normal;

    /// <summary>Full paths of files to attach.</summary>
    public IReadOnlyList<string> Attachments { get; set; } = [];
}

/// <summary>The details of a meeting to send invitations for.</summary>
/// <remarks>A meeting is an appointment with attendees, so it extends one.</remarks>
public sealed class MeetingEntry : CalendarEntry
{
    /// <summary>Semicolon-separated required attendees.</summary>
    public string RequiredAttendees { get; set; } = string.Empty;

    /// <summary>Semicolon-separated optional attendees.</summary>
    public string OptionalAttendees { get; set; } = string.Empty;

    /// <summary>Whether attendees are asked to respond.</summary>
    public bool ResponseRequested { get; set; }
}
