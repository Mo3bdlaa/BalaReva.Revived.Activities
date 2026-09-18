using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.External;

/// <summary>Copies a range to the clipboard.</summary>
[DisplayName("Copy Data")]
[Description("Copies a range of cells to the Windows clipboard.")]
public sealed class CopyData : ExcelCore
{
    /// <summary>Range to copy.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Copy Range")]
    [Description("Range to copy, for example A1:D20.")]
    public InArgument<string> CopyRange { get; set; } = null!;

    /// <summary>True when the range was copied.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the range was copied.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.CopyData(sheetName, Require(context, CopyRange, nameof(CopyRange)));

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
