using System.Activities;
using System.ComponentModel;

namespace BalaReva.Printer;

/// <summary>Runs a queue operation against the default printer.</summary>
/// <remarks>
/// This one takes no printer name, matching the published package, so it always acts
/// on the machine default. Use <see cref="PausePrinter"/>, <see cref="ResumePrinter"/>
/// or <see cref="ClearPrinterQueue"/> to target a named printer.
/// </remarks>
[DisplayName("Printer Command")]
[Description("Runs a queue operation against the default printer.")]
public sealed class PrinterCommand : BaseActivity
{
    /// <summary>Operation to run.</summary>
    [Category("Input")]
    [DisplayName("Command")]
    [Description("The queue operation to run on the default printer.")]
    public CommandEnum Command { get; set; } = CommandEnum.Refresh;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPrinterService service) =>
        service.RunCommand(Command);
}
