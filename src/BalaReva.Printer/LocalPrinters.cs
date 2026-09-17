using System.Activities;
using System.ComponentModel;

namespace BalaReva.Printer;

/// <summary>Lists the print queues installed on this machine.</summary>
[DisplayName("Local Printers")]
[Description("Lists the names of the print queues installed locally.")]
public sealed class LocalPrinters : BaseActivity
{
    /// <summary>Names of the local print queues.</summary>
    [Category("Output")]
    [DisplayName("Printers")]
    [Description("Names of the locally installed print queues.")]
    public OutArgument<string[]> Printers { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPrinterService service) =>
        Printers.Set(context, service.GetLocalPrinters());
}
