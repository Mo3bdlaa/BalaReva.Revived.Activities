using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook.Meetings;

/// <summary>Sends a meeting invitation from Outlook.</summary>
/// <remarks>
/// Unlike <c>NewAppointment</c>, which saves to the calendar, this sends invitations to
/// the attendees as soon as it runs.
/// </remarks>
[DisplayName("New Meeting")]
[Description("Sends a meeting invitation from Outlook.")]
public sealed class NewMeeting : BaseActivity
{
    /// <summary>Outlook account to send from. Leave empty for the default.</summary>
    [Category("Input")]
    [DisplayName("Account")]
    [Description("Outlook account to send from. Leave empty for the default account.")]
    public InArgument<string> Account { get; set; } = null!;

    /// <summary>Subject line.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Subject")]
    [Description("Subject line of the meeting.")]
    public InArgument<string> Subject { get; set; } = null!;

    /// <summary>Body text.</summary>
    [Category("Input")]
    [DisplayName("Body")]
    [Description("Body text of the invitation.")]
    public InArgument<string> Body { get; set; } = null!;

    /// <summary>Where it takes place.</summary>
    [Category("Input")]
    [DisplayName("Appointment Location")]
    [Description("Where the meeting takes place.")]
    public InArgument<string> AppointmentLocation { get; set; } = null!;

    /// <summary>When it starts.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Start Date")]
    [Description("When the meeting starts. The time is ignored for an all-day event.")]
    public InArgument<DateTime> StartDate { get; set; } = null!;

    /// <summary>When it ends.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("End Date")]
    [Description("When the meeting ends. The time is ignored for an all-day event.")]
    public InArgument<DateTime> EndDate { get; set; } = null!;

    /// <summary>True for an all-day entry.</summary>
    [Category("Input")]
    [DisplayName("All Day Event")]
    [Description("True for an all-day entry, in which case the times are ignored.")]
    public InArgument<bool> AllDayEvent { get; set; } = null!;

    /// <summary>Minutes before the start to show a reminder.</summary>
    [Category("Input")]
    [DisplayName("Reminder Minutes")]
    [Description("Minutes before the start to show a reminder. Zero means no reminder.")]
    public InArgument<int> ReminderMinutes { get; set; } = null!;

    /// <summary>Full paths of files to attach.</summary>
    [Category("Input")]
    [DisplayName("Attachments")]
    [Description("Full paths of files to attach.")]
    public InArgument<List<string>> Attachments { get; set; } = null!;

    /// <summary>Attendees who must come.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Required Attendees")]
    [Description("Semicolon-separated addresses of attendees who must come.")]
    public InArgument<string> RequiredAttendees { get; set; } = null!;

    /// <summary>Attendees who may come.</summary>
    [Category("Input")]
    [DisplayName("Optional Attendees")]
    [Description("Semicolon-separated addresses of attendees who may come.")]
    public InArgument<string> OptionalAttendees { get; set; } = null!;

    /// <summary>Whether attendees are asked to respond.</summary>
    [Category("Input")]
    [DisplayName("Response Requested")]
    [Description("Whether attendees are asked to accept or decline.")]
    public InArgument<bool> ResponseRequested { get; set; } = null!;

    /// <summary>How it shows on the calendar.</summary>
    [Category("Input")]
    [DisplayName("Busy Status")]
    [Description("How the meeting shows on the calendar.")]
    public BusyStatusEnum BusyStatus { get; set; } = BusyStatusEnum.Busy;

    /// <summary>Importance flag.</summary>
    [Category("Input")]
    [DisplayName("Importance")]
    [Description("Importance flag on the invitation.")]
    public ImportanceEnum Importance { get; set; } = ImportanceEnum.Normal;

    /// <summary>True when the invitation was sent.</summary>
    [Category("Output")]
    [DisplayName("Added Result")]
    [Description("True when the invitation was sent.")]
    public OutArgument<bool> AddedResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service)
    {
        var entry = new MeetingEntry
        {
            Account = Account?.Get(context) ?? string.Empty,
            Subject = Require(context, Subject, nameof(Subject)),
            Body = Body?.Get(context) ?? string.Empty,
            Location = AppointmentLocation?.Get(context) ?? string.Empty,
            Start = StartDate.Get(context),
            End = EndDate.Get(context),
            AllDayEvent = AllDayEvent?.Get(context) ?? false,
            ReminderMinutes = ReminderMinutes?.Get(context) ?? 0,
            BusyStatus = BusyStatus,
            Importance = Importance,
            Attachments = Attachments?.Get(context) ?? [],
            RequiredAttendees = Require(context, RequiredAttendees, nameof(RequiredAttendees)),
            OptionalAttendees = OptionalAttendees?.Get(context) ?? string.Empty,
            ResponseRequested = ResponseRequested?.Get(context) ?? false,
        };

        if (!entry.AllDayEvent && entry.End < entry.Start)
            throw new ArgumentException("End Date must not be before Start Date.", nameof(EndDate));

        AddedResult.Set(context, service.AddMeeting(entry));
    }
}
