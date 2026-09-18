using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.External;

/// <summary>Clears a range of cells.</summary>
[DisplayName("Delete Data")]
[Description("Clears the contents and formatting of a range of cells.")]
public sealed class DeleteData : ExcelCore
{
    /// <summary>Range to clear.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Delete Range")]
    [Description("Range to clear, for example A1:D20.")]
    public InArgument<string> DeleteRange { get; set; } = null!;

    /// <summary>True when the range was cleared.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the range was cleared.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.DeleteData(sheetName, Require(context, DeleteRange, nameof(DeleteRange)));

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
