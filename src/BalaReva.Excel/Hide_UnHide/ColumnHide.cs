using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.Hide_UnHide;

/// <summary>Hides or shows named columns.</summary>
[DisplayName("Column Hide")]
[Description("Hides or shows columns by their letter.")]
public sealed class ColumnHide : ExcelCore
{
    /// <summary>Columns to hide or show, by letter.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Column Names")]
    [Description("Columns to hide or show, by letter, for example A or C.")]
    public InArgument<string[]> ColumnNames { get; set; } = null!;

    /// <summary>Whether to hide or show them.</summary>
    [Category("Input")]
    [DisplayName("Hidden Type")]
    [Description("Whether to hide or show the columns.")]
    public HideEnum HiddenType { get; set; } = HideEnum.Hide;

    /// <summary>True when the columns were set.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the columns were hidden or shown.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.HideColumns(sheetName, ColumnNames?.Get(context) ?? [], HiddenType == HideEnum.Hide);

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
