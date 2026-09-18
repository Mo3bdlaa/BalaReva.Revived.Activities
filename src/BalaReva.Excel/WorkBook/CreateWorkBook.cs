using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.WorkBook;

/// <summary>Creates a new workbook holding one named sheet.</summary>
[DisplayName("CreateWorkBook")]
[Description("Creates a new workbook holding one named sheet.")]
public sealed class CreateWorkBook : ExcelCore
{
    /// <summary>True when the sheet was created.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the sheet was created.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.CreateWorkBook(sheetName);

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
