using System.Activities;
using System.ComponentModel;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.Documents;

/// <summary>Prints the document.</summary>
[DisplayName("Print Document")]
[Description("Prints the document on a named printer.")]
public sealed class PrintDocument : BaseNativeChild
{
    /// <summary>Printer to print on. Empty uses the default.</summary>
    [Category("Input")]
    [DisplayName("Printer Name")]
    [Description("Printer to print on. Leave empty for the default printer.")]
    public InArgument<string> PrinterName { get; set; } = null!;

    /// <summary>How many copies.</summary>
    [Category("Input")]
    [DisplayName("No Of Copies")]
    [Description("How many copies to print. Less than one is treated as one.")]
    public InArgument<int> NoOfCopies { get; set; } = null!;

    /// <summary>Page orientation.</summary>
    [Category("Input")]
    [DisplayName("Orientation")]
    [Description("Page orientation. Select leaves the document's own setting alone.")]
    public EnumOrientation Orientation { get; set; } = EnumOrientation.Select;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.Print(
            PrinterName?.Get(context) ?? string.Empty,
            NoOfCopies?.Get(context) ?? 1,
            Orientation);
}
