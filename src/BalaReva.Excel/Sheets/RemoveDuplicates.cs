using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.Sheets;

/// <summary>Removes duplicate rows from a range.</summary>
[DisplayName("Remove Duplicates")]
[Description("Removes duplicate rows from a range, comparing the given columns.")]
public sealed class RemoveDuplicates : ExcelCore
{
    /// <summary>Range to deduplicate. Empty uses the whole used range.</summary>
    [Category("Input")]
    [DisplayName("Cell Range")]
    [Description("Range to deduplicate. Empty uses the whole used range.")]
    public InArgument<string> CellRange { get; set; } = null!;

    /// <summary>Column numbers to compare. Empty compares every column.</summary>
    [Category("Input")]
    [DisplayName("Columns")]
    [Description("Column numbers within the range to compare. Empty compares them all.")]
    public InArgument<object[]> Columns { get; set; } = null!;

    /// <summary>Whether the first row holds column names.</summary>
    [Category("Input")]
    [DisplayName("Has Header")]
    [Description("Treat the first row as a header and leave it in place.")]
    public InArgument<bool> HasHeader { get; set; } = null!;

    /// <summary>True when the duplicates were removed.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the duplicates were removed.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.RemoveDuplicates(
            sheetName,
            CellRange?.Get(context) ?? string.Empty,
            Columns?.Get(context) ?? [],
            HasHeader?.Get(context) ?? false);

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
