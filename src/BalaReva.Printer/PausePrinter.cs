using System.Activities;
using System.ComponentModel;
using BalaReva.Printer.Enums;

namespace BalaReva.Printer;

/// <summary>Pauses a printer's queue.</summary>
[DisplayName("Pause Printer")]
[Description("Pauses the named printer's queue.")]
public sealed class PausePrinter : BaseActivity
{
    /// <summary>Printer to pause.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Printer Name")]
    [Description("Name of the printer whose queue to pause.")]
    public InArgument<string> PrinterName { get; set; } = null!;

    /// <summary>Access to open the queue with.</summary>
    [Category("Input")]
    [DisplayName("Desired Access")]
    [Description("Access to open the queue with. None lets the activity pick what the operation needs.")]
    public AccessRightsEnum DesiredAccess { get; set; } = AccessRightsEnum.None;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPrinterService service) =>
        service.Pause(RequirePrinterName(context, PrinterName), DesiredAccess);
}
