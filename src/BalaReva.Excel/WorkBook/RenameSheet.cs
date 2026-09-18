using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.WorkBook;

/// <summary>Renames a sheet.</summary>
[DisplayName("Rename Sheet")]
[Description("Renames a sheet.")]
public sealed class RenameSheet : ExcelCore
{
    /// <summary>New name for the sheet.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("New Sheet Name")]
    [Description("New name for the sheet.")]
    public InArgument<string> NewSheetName { get; set; } = null!;

    /// <summary>True when the sheet was renamed.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the sheet was renamed.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.RenameSheet(sheetName, Require(context, NewSheetName, nameof(NewSheetName)));

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
