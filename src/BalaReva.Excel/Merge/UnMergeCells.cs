using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;
using BalaReva.Excel.Enums;

namespace BalaReva.Excel.Merge;

/// <summary>Splits a merged range back into cells.</summary>
[DisplayName("Un Merge Cells")]
[Description("Splits a merged range back into individual cells.")]
public sealed class UnMergeCells : ExcelCore
{
    /// <summary>Range to unmerge.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Un Merge Range")]
    [Description("Range to unmerge, for example A1:C1.")]
    public InArgument<string> UnMergeRange { get; set; } = null!;

    /// <summary>Horizontal alignment to leave behind.</summary>
    [Category("Input")]
    [DisplayName("Horizontal Alignment")]
    [Description("Horizontal alignment applied to the cells afterwards.")]
    public AlignmentEnum HorizontalAlignment { get; set; } = AlignmentEnum.General;

    /// <summary>Vertical alignment to leave behind.</summary>
    [Category("Input")]
    [DisplayName("Vertical Alignment")]
    [Description("Vertical alignment applied to the cells afterwards.")]
    public AlignmentEnum VerticalAlignment { get; set; } = AlignmentEnum.Center;

    /// <summary>True when the range was unmerged.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the range was unmerged.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.UnMergeCells(
            sheetName,
            Require(context, UnMergeRange, nameof(UnMergeRange)),
            HorizontalAlignment,
            VerticalAlignment);

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
