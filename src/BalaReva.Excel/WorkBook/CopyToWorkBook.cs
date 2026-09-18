using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.WorkBook;

/// <summary>Copies a sheet within the same workbook.</summary>
[DisplayName("Copy To Work Book")]
[Description("Copies a sheet to the end of the same workbook.")]
public sealed class CopyToWorkBook : ExcelCore
{
    /// <summary>Name for the copy.</summary>
    [Category("Input")]
    [DisplayName("New Sheet Name")]
    [Description("Name for the copied sheet. Empty lets Excel name it.")]
    public InArgument<string> NewSheetName { get; set; } = null!;

    /// <summary>True when the sheet was copied.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the sheet was copied.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.CopyToWorkBook(sheetName, NewSheetName?.Get(context) ?? string.Empty);

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
