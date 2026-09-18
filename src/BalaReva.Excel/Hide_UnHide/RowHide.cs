using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.Hide_UnHide;

/// <summary>Hides or shows numbered rows.</summary>
[DisplayName("Row Hide")]
[Description("Hides or shows rows by their number.")]
public sealed class RowHide : ExcelCore
{
    /// <summary>Rows to hide or show. Rows are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Row Numbers")]
    [Description("Rows to hide or show, numbered from 1.")]
    public InArgument<int[]> RowNumbers { get; set; } = null!;

    /// <summary>Whether to hide or show them.</summary>
    [Category("Input")]
    [DisplayName("Hidden Type")]
    [Description("Whether to hide or show the rows.")]
    public HideEnum HiddenType { get; set; } = HideEnum.Hide;

    /// <summary>True when the rows were set.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the rows were hidden or shown.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.HideRows(sheetName, RowNumbers?.Get(context) ?? [], HiddenType == HideEnum.Hide);

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
