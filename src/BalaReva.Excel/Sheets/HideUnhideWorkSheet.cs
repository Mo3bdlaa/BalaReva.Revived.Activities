using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.Sheets;

/// <summary>Shows or hides a sheet.</summary>
[DisplayName("Hide Unhide Work Sheet")]
[Description("Shows or hides a sheet.")]
public sealed class HideUnhideWorkSheet : ExcelCore
{
    /// <summary>Whether the sheet is visible.</summary>
    [Category("Input")]
    [DisplayName("Sheet Visibility")]
    [Description("Whether the sheet is visible or hidden.")]
    public XlSheetVisibility SheetVisibility { get; set; } = XlSheetVisibility.Visible;

    /// <summary>True when the visibility was set.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the visibility was set.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.SetSheetVisibility(sheetName, SheetVisibility);

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
