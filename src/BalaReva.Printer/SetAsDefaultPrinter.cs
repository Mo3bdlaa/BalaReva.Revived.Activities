using System.Activities;
using System.ComponentModel;

namespace BalaReva.Printer;

/// <summary>Makes a printer the machine default.</summary>
[DisplayName("Set As Default Printer")]
[Description("Makes the named printer the machine's default.")]
public sealed class SetAsDefaultPrinter : BaseActivity
{
    /// <summary>Printer to make the default.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Printer Name")]
    [Description("Name of the printer to make the default.")]
    public InArgument<string> PrinterName { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPrinterService service) =>
        service.SetDefaultPrinter(RequirePrinterName(context, PrinterName));
}
