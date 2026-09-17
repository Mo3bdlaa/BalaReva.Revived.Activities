using System.Activities;
using System.ComponentModel;

namespace BalaReva.Printer;

/// <summary>Reads the state of the machine's default printer.</summary>
[DisplayName("Default Printer Status")]
[Description("Reads a snapshot of the default printer's state.")]
public sealed class DefaultPrinterStatus : BaseActivity
{
    /// <summary>A snapshot of the default printer's state.</summary>
    [Category("Output")]
    [DisplayName("Printer Result")]
    [Description("A snapshot of the default printer's state.")]
    public OutArgument<PrinterStatus> PrinterResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPrinterService service) =>
        PrinterResult.Set(context, service.GetDefaultPrinterStatus());
}
