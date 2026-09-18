using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Copies the formatting of one range onto another.</summary>
[DisplayName("Format Painter")]
[Description("Copies the formatting of one range onto another.")]
public sealed class FormatPainter : BaseActivity
{
    /// <summary>Range to copy the formatting from.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell Range")]
    [Description("Range to copy the formatting from.")]
    public InArgument<string> CellRange { get; set; } = null!;

    /// <summary>Range to paste the formatting onto.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Destination Cell Range")]
    [Description("Range to paste the formatting onto.")]
    public InArgument<string> DestinationCellRange { get; set; } = null!;

    /// <summary>Sheet to paste onto. Empty means the same sheet.</summary>
    [Category("Input")]
    [DisplayName("Destination Sheet")]
    [Description("Sheet to paste onto. Empty means the same sheet.")]
    public InArgument<string> DestinationSheet { get; set; } = null!;

    /// <summary>True when the formatting was copied.</summary>
    [Category("Output")]
    [DisplayName("Output")]
    [Description("True when the formatting was copied.")]
    public OutArgument<bool> Output { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => Output.Set(context, workbook.FormatPainter(
            sheetName,
            Require(context, CellRange, nameof(CellRange)),
            DestinationSheet?.Get(context) ?? string.Empty,
            Require(context, DestinationCellRange, nameof(DestinationCellRange))));
}
