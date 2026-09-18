using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.WorkBook;

/// <summary>Adds a sheet with the given name.</summary>
[DisplayName("AddSheet")]
[Description("Adds a sheet with the given name.")]
public sealed class AddSheet : ExcelCore
{
    /// <summary>True when the sheet was added.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the sheet was added.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.AddSheet(sheetName);

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
