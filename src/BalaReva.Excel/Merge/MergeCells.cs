using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;
using BalaReva.Excel.Enums;

namespace BalaReva.Excel.Merge;

/// <summary>Merges a range into one cell.</summary>
[DisplayName("Merge Cells")]
[Description("Merges a range of cells into one and sets its text and alignment.")]
public sealed class MergeCells : ExcelCore
{
    /// <summary>Range to merge.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Merge Range")]
    [Description("Range to merge, for example A1:C1.")]
    public InArgument<string> MergeRange { get; set; } = null!;

    /// <summary>Text for the merged cell.</summary>
    [Category("Input")]
    [DisplayName("Cell Text")]
    [Description("Text for the merged cell. Empty leaves the existing value.")]
    public InArgument<string> CellText { get; set; } = null!;

    /// <summary>Horizontal alignment.</summary>
    [Category("Input")]
    [DisplayName("Horizontal Alignment")]
    [Description("Horizontal alignment of the merged cell.")]
    public AlignmentEnum HorizontalAlignment { get; set; } = AlignmentEnum.General;

    /// <summary>Vertical alignment.</summary>
    [Category("Input")]
    [DisplayName("Vertical Alignment")]
    [Description("Vertical alignment of the merged cell.")]
    public AlignmentEnum VerticalAlignment { get; set; } = AlignmentEnum.Center;

    /// <summary>True when the range was merged.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the range was merged.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.MergeCells(
            sheetName,
            Require(context, MergeRange, nameof(MergeRange)),
            CellText?.Get(context) ?? string.Empty,
            HorizontalAlignment,
            VerticalAlignment);

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
