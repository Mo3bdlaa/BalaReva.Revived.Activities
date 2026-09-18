using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.Sheets;

/// <summary>Autofits or sets the height of rows.</summary>
[DisplayName("Auto Fit Rows")]
[Description("Autofits or sets the height of rows.")]
public sealed class AutoFitRows : ExcelCore
{
    /// <summary>Rows to size. Empty means every used row.</summary>
    [Category("Input")]
    [DisplayName("Rows Range")]
    [Description("Rows to size, numbered from 1. Empty sizes every used row.")]
    public InArgument<int[]> RowsRange { get; set; } = null!;

    /// <summary>Whether to autofit rather than use a fixed height.</summary>
    [Category("Input")]
    [DisplayName("Auto Fit")]
    [Description("Autofit to the content. When false, Row Height is used instead.")]
    public InArgument<bool> AutoFit { get; set; } = null!;

    /// <summary>Fixed height, when not autofitting.</summary>
    [Category("Input")]
    [DisplayName("Row Height")]
    [Description("Fixed height in points, used when Auto Fit is false.")]
    public InArgument<double> RowHeight { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.AutoFitRows(
            sheetName,
            RowsRange?.Get(context) ?? [],
            AutoFit?.Get(context) ?? true,
            RowHeight?.Get(context) ?? 0);

    /// <summary>True when the activity completed.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the activity completed without error.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
