using System.Activities;
using System.ComponentModel;

namespace BalaReva.Printer;

/// <summary>Lists the print queues reached through a print server connection.</summary>
[DisplayName("Network Printers")]
[Description("Lists the names of print queues connected from a print server.")]
public sealed class NetworkPrinters : BaseActivity
{
    /// <summary>Names of the connected network print queues.</summary>
    [Category("Output")]
    [DisplayName("Printers")]
    [Description("Names of the print queues connected from a print server.")]
    public OutArgument<string[]> Printers { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPrinterService service) =>
        Printers.Set(context, service.GetNetworkPrinters());
}
