using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;
using BalaReva.Excel.Utilities;

namespace BalaReva.Excel.WorkBook;

/// <summary>Exports the workbook, or part of it, to PDF or XPS.</summary>
[DisplayName("Export Work Book")]
[Description("Exports the workbook, or a range of one sheet, to PDF or XPS.")]
public sealed class ExportWorkBook : ExcelCore
{
    /// <summary>Path to write the export to.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Export Path")]
    [Description("Full path to write the exported file to.")]
    public InArgument<string> ExportPath { get; set; } = null!;

    /// <summary>Range to export.</summary>
    [Category("Input")]
    [DisplayName("Cell Range")]
    [Description("Range to export. Empty exports the whole workbook.")]
    public InArgument<string> CellRange { get; set; } = null!;

    /// <summary>Format to export in.</summary>
    [Category("Input")]
    [DisplayName("Format Type")]
    [Description("Whether to export as PDF or XPS.")]
    public FixedFormatTypeEnum FormatType { get; set; } = FixedFormatTypeEnum.PDF;

    /// <summary>True when the export was written.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the export was written.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.ExportWorkBook(
            sheetName,
            CellRange?.Get(context) ?? string.Empty,
            Require(context, ExportPath, nameof(ExportPath)),
            FormatType);

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
