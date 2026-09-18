using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;
using BalaReva.Excel.Enums;

namespace BalaReva.Excel.Sheets;

/// <summary>Sets alignment and text control on ranges.</summary>
/// <remarks>
/// Every enum here carries a <c>Select</c> member meaning "leave this alone", so an
/// untouched property does not overwrite the sheet's existing formatting.
/// </remarks>
[DisplayName("Format Cells")]
[Description("Sets alignment, wrapping and orientation on one or more ranges.")]
public sealed class FormatCells : ExcelCore
{
    /// <summary>Ranges to format.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell Range")]
    [Description("Ranges to format, for example A1:C10.")]
    public InArgument<string[]> CellRange { get; set; } = null!;

    /// <summary>Horizontal alignment.</summary>
    [Category("Input")]
    [DisplayName("Horizontal")]
    [Description("Horizontal alignment. Select leaves it alone.")]
    public FormatHorizontalEnum Horizontal { get; set; } = FormatHorizontalEnum.Select;

    /// <summary>Vertical alignment.</summary>
    [Category("Input")]
    [DisplayName("Verticle")]
    [Description("Vertical alignment. Select leaves it alone.")]
    public FormatVerticleEnum Verticle { get; set; } = FormatVerticleEnum.Select;

    /// <summary>Wrap text within the cell.</summary>
    [Category("Input")]
    [DisplayName("Wrap Text")]
    [Description("Wrap text within the cell. Select leaves it alone.")]
    public FormatTextControlEnum WrapText { get; set; } = FormatTextControlEnum.Select;

    /// <summary>Shrink text to fit.</summary>
    [Category("Input")]
    [DisplayName("Shrink Fit")]
    [Description("Shrink the text to fit the cell. Select leaves it alone.")]
    public FormatTextControlEnum ShrinkFit { get; set; } = FormatTextControlEnum.Select;

    /// <summary>Merge the range.</summary>
    [Category("Input")]
    [DisplayName("Merge Cells")]
    [Description("Merge the range into one cell. Select leaves it alone.")]
    public FormatTextControlEnum MergeCells { get; set; } = FormatTextControlEnum.Select;

    /// <summary>Reading order.</summary>
    [Category("Input")]
    [DisplayName("Text Direction")]
    [Description("Reading order. Select leaves it alone.")]
    public FormatTextDirection TextDirection { get; set; } = FormatTextDirection.Select;

    /// <summary>Preset text orientation.</summary>
    [Category("Input")]
    [DisplayName("Text Orientation")]
    [Description("Preset text orientation. Select falls back to Orientation Deg.")]
    public TextOrientationEumn TextOrientation { get; set; } = TextOrientationEumn.Select;

    /// <summary>Rotation in degrees.</summary>
    [Category("Input")]
    [DisplayName("Orientation Deg")]
    [Description("Rotation in degrees, used when Text Orientation is Select.")]
    public InArgument<int> OrientationDeg { get; set; } = null!;

    /// <summary>Indent level.</summary>
    [Category("Input")]
    [DisplayName("Indent")]
    [Description("Indent level. Zero leaves it alone.")]
    public InArgument<int> Indent { get; set; } = null!;

    /// <summary>True when the ranges were formatted.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the ranges were formatted.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.FormatCells(sheetName, CellRange?.Get(context) ?? [], new CellFormatRequest
        {
            Horizontal = Horizontal,
            Verticle = Verticle,
            WrapText = WrapText,
            ShrinkFit = ShrinkFit,
            MergeCells = MergeCells,
            TextDirection = TextDirection,
            TextOrientation = TextOrientation,
            OrientationDeg = OrientationDeg?.Get(context) ?? 0,
            Indent = Indent?.Get(context) ?? 0,
        });

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
