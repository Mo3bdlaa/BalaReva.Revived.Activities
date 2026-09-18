using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.Sheets;

/// <summary>Removes the hyperlinks from a range.</summary>
[DisplayName("Remove Hyperlink")]
[Description("Removes the hyperlinks from a range.")]
public sealed class RemoveHyperlink : ExcelCore
{
    /// <summary>Range to strip. Empty strips the whole used range.</summary>
    [Category("Input")]
    [DisplayName("Cell Range")]
    [Description("Range to strip hyperlinks from. Empty strips the whole used range.")]
    public InArgument<string> CellRange { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.RemoveHyperlink(sheetName, CellRange?.Get(context) ?? string.Empty);

    /// <summary>True when the activity completed.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the activity completed without error.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
