using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.Sheets;

/// <summary>Deletes a sheet from the workbook.</summary>
[DisplayName("Delete Sheet")]
[Description("Deletes a sheet from the workbook.")]
public sealed class DeleteSheet : ExcelCore
{
    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.DeleteSheet(sheetName);

    /// <summary>True when the activity completed.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the activity completed without error.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
