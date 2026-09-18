using System.Activities;
using System.ComponentModel;
using System.Data;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.Sheets;

/// <summary>Turns column values into hyperlinks.</summary>
[DisplayName("Hyperlink Add")]
[Description("Turns the values in named columns into hyperlinks.")]
public sealed class HyperlinkAdd : ExcelCore
{
    /// <summary>Table holding the addresses to link to.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Input Table")]
    [Description("Table holding the addresses, matching the sheet's rows.")]
    public InArgument<DataTable> InputTable { get; set; } = null!;

    /// <summary>Columns to turn into links.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Column Names")]
    [Description("Columns whose values become hyperlinks.")]
    public InArgument<string[]> ColumnNames { get; set; } = null!;

    /// <summary>Whether to replace the displayed text with the address.</summary>
    [Category("Input")]
    [DisplayName("Text Overwirte")]
    [Description("Replace the cell's displayed text with the address.")]
    public InArgument<bool> TextOverwirte { get; set; } = null!;

    /// <summary>True when the hyperlinks were added.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the hyperlinks were added.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
    {
        var input = InputTable?.Get(context)
            ?? throw new ArgumentException("InputTable is required.", nameof(InputTable));

        workbook.AddHyperlinks(
            sheetName,
            input,
            ColumnNames?.Get(context) ?? [],
            TextOverwirte?.Get(context) ?? false);
    }

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
