using System.Activities;
using System.ComponentModel;
using BalaReva.Printer.Enums;

namespace BalaReva.Printer;

/// <summary>Cancels every job waiting on a printer.</summary>
[DisplayName("Clear Printer Queue")]
[Description("Cancels every job waiting on the named printer.")]
public sealed class ClearPrinterQueue : BaseActivity
{
    /// <summary>Printer whose queue to clear.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Printer Name")]
    [Description("Name of the printer whose queue to clear.")]
    public InArgument<string> PrinterName { get; set; } = null!;

    /// <summary>Access to open the queue with.</summary>
    [Category("Input")]
    [DisplayName("Desired Access")]
    [Description("Access to open the queue with. None lets the activity pick what the operation needs.")]
    public AccessRightsEnum DesiredAccess { get; set; } = AccessRightsEnum.None;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPrinterService service) =>
        service.ClearQueue(RequirePrinterName(context, PrinterName), DesiredAccess);
}
