using System.Activities;
using System.ComponentModel;

namespace BalaReva.Printer;

/// <summary>Reads the name of the machine's default printer.</summary>
[DisplayName("Get Default Printer")]
[Description("Reads the name of the machine's default printer.")]
public sealed class GetDefaultPrinter : BaseActivity
{
    /// <summary>Name of the default printer.</summary>
    [Category("Output")]
    [DisplayName("Output")]
    [Description("Name of the default printer.")]
    public OutArgument<string> Output { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPrinterService service) =>
        Output.Set(context, service.GetDefaultPrinterName());
}
