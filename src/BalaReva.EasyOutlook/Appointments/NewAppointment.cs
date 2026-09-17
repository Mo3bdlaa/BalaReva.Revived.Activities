using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook.Appointments;

/// <summary>Adds an appointment to the Outlook calendar.</summary>
[DisplayName("New Appointment")]
[Description("Adds an appointment to the Outlook calendar.")]
public sealed class NewAppointment : BaseActivity
{
    /// <summary>Outlook account whose calendar to add to. Leave empty for the default.</summary>
    [Category("Input")]
    [DisplayName("Account")]
    [Description("Outlook account whose calendar to add to. Leave empty for the default account.")]
    public InArgument<string> Account { get; set; } = null!;

    /// <summary>Subject line.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Subject")]
    [Description("Subject line of the appointment.")]
    public InArgument<string> Subject { get; set; } = null!;

    /// <summary>Body text.</summary>
    [Category("Input")]
    [DisplayName("Body")]
    [Description("Body text of the appointment.")]
    public InArgument<string> Body { get; set; } = null!;

    /// <summary>Where it takes place.</summary>
    [Category("Input")]
    [DisplayName("Appointment Location")]
    [Description("Where the appointment takes place.")]
    public InArgument<string> AppointmentLocation { get; set; } = null!;

    /// <summary>When it starts.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Start Date")]
    [Description("When the appointment starts. The time is ignored for an all-day event.")]
    public InArgument<DateTime> StartDate { get; set; } = null!;

    /// <summary>When it ends.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("End Date")]
    [Description("When the appointment ends. The time is ignored for an all-day event.")]
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

    /// <summary>How it shows on the calendar.</summary>
    [Category("Input")]
    [DisplayName("Busy Status")]
    [Description("How the appointment shows on the calendar.")]
    public BusyStatusEnum BusyStatus { get; set; } = BusyStatusEnum.Busy;

    /// <summary>Importance flag.</summary>
    [Category("Input")]
    [DisplayName("Importance")]
    [Description("Importance flag on the appointment.")]
    public ImportanceEnum Importance { get; set; } = ImportanceEnum.Normal;

    /// <summary>True when the appointment was saved.</summary>
    [Category("Output")]
    [DisplayName("Added Result")]
    [Description("True when the appointment was saved.")]
    public OutArgument<bool> AddedResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service)
    {
        var entry = new CalendarEntry();
        entry.Account = Account?.Get(context) ?? string.Empty;
        entry.Subject = Require(context, Subject, nameof(Subject));
        entry.Body = Body?.Get(context) ?? string.Empty;
        entry.Location = AppointmentLocation?.Get(context) ?? string.Empty;
        entry.Start = StartDate.Get(context);
        entry.End = EndDate.Get(context);
        entry.AllDayEvent = AllDayEvent?.Get(context) ?? false;
        entry.ReminderMinutes = ReminderMinutes?.Get(context) ?? 0;
        entry.BusyStatus = BusyStatus;
        entry.Importance = Importance;
        entry.Attachments = Attachments?.Get(context) ?? [];

        if (!entry.AllDayEvent && entry.End < entry.Start)
        {
            throw new ArgumentException(
                "End Date must not be before Start Date.", nameof(EndDate));
        }

        AddedResult.Set(context, service.AddAppointment(entry));
    }
}
