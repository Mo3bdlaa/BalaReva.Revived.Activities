using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;
using BalaReva.Excel.Enums;

namespace BalaReva.Excel.Sheets;

/// <summary>Applies a table style to an existing table.</summary>
[DisplayName("Set Table Format")]
[Description("Applies a built-in or named table style to an existing table.")]
public sealed class SetTableFormat : ExcelCore
{
    /// <summary>Which table. Tables are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Table Index")]
    [Description("Which table on the sheet. Tables are numbered from 1.")]
    public InArgument<int> TableIndex { get; set; } = null!;

    /// <summary>Built-in style to apply.</summary>
    [Category("Input")]
    [DisplayName("Table Format Style")]
    [Description("Built-in style to apply. Ignored when Custom Style is set.")]
    public TableFormatEnum TableFormatStyle { get; set; } = TableFormatEnum.TableStyleLight1;

    /// <summary>Named style to apply instead.</summary>
    [Category("Input")]
    [DisplayName("Custom Style")]
    [Description("Name of a table style to apply. Takes precedence over Table Format Style.")]
    public InArgument<string> CustomStyle { get; set; } = null!;

    /// <summary>True when the style was applied.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the style was applied.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.SetTableFormat(
            sheetName,
            TableIndex.Get(context),
            TableFormatStyle,
            CustomStyle?.Get(context) ?? string.Empty);

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
