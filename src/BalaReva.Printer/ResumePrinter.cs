using System.Activities;
using System.ComponentModel;
using BalaReva.Printer.Enums;

namespace BalaReva.Printer;

/// <summary>Resumes a paused printer queue.</summary>
[DisplayName("Resume Printer")]
[Description("Resumes the named printer's paused queue.")]
public sealed class ResumePrinter : BaseActivity
{
    /// <summary>Printer to resume.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Printer Name")]
    [Description("Name of the printer whose queue to resume.")]
    public InArgument<string> PrinterName { get; set; } = null!;

    /// <summary>Access to open the queue with.</summary>
    [Category("Input")]
    [DisplayName("Desired Access")]
    [Description("Access to open the queue with. None lets the activity pick what the operation needs.")]
    public AccessRightsEnum DesiredAccess { get; set; } = AccessRightsEnum.None;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPrinterService service) =>
        service.Resume(RequirePrinterName(context, PrinterName), DesiredAccess);
}
