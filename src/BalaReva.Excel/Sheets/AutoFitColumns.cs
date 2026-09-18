using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.Sheets;

/// <summary>Autofits or sets the width of columns.</summary>
[DisplayName("Auto Fit Columns")]
[Description("Autofits or sets the width of columns.")]
public sealed class AutoFitColumns : ExcelCore
{
    /// <summary>Columns to size, by letter. Empty means every used column.</summary>
    [Category("Input")]
    [DisplayName("Columns Range")]
    [Description("Columns to size, by letter. Empty sizes every used column.")]
    public InArgument<string[]> ColumnsRange { get; set; } = null!;

    /// <summary>Whether to autofit rather than use a fixed width.</summary>
    [Category("Input")]
    [DisplayName("Auto Fit")]
    [Description("Autofit to the content. When false, Column Width is used instead.")]
    public InArgument<bool> AutoFit { get; set; } = null!;

    /// <summary>Fixed width, when not autofitting.</summary>
    [Category("Input")]
    [DisplayName("Column Width")]
    [Description("Fixed width in characters, used when Auto Fit is false.")]
    public InArgument<double> ColumnWidth { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.AutoFitColumns(
            sheetName,
            ColumnsRange?.Get(context) ?? [],
            AutoFit?.Get(context) ?? true,
            ColumnWidth?.Get(context) ?? 0);

    /// <summary>True when the activity completed.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the activity completed without error.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
