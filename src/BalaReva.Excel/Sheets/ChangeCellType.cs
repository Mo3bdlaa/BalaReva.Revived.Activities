using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.Sheets;

/// <summary>Sets a cell's number format.</summary>
[DisplayName("Change Cell Type")]
[Description("Sets a cell's number format.")]
public sealed class ChangeCellType : ExcelCore
{
    /// <summary>Cell to format.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell")]
    [Description("Cell to format, for example B4.")]
    public InArgument<string> Cell { get; set; } = null!;

    /// <summary>Excel number format string.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell Format")]
    [Description("Excel number format string, for example 0.00 or dd/mm/yyyy.")]
    public InArgument<string> CellFormat { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.ChangeCellType(
            sheetName,
            Require(context, Cell, nameof(Cell)),
            Require(context, CellFormat, nameof(CellFormat)));

    /// <summary>True when the activity completed.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the activity completed without error.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
