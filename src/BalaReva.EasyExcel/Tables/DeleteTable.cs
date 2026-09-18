using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Tables;

/// <summary>Deletes a table, leaving its cells behind.</summary>
[DisplayName("Delete Table")]
[Description("Deletes a table, leaving its cells behind.")]
public sealed class DeleteTable : BaseActivity
{
    /// <summary>Which table on the sheet, numbered from 1.</summary>
    [Category("Input")]
    [DisplayName("Table Index")]
    [Description("Which table on the sheet, numbered from 1.")]
    public InArgument<int> TableIndex { get; set; } = null!;

    /// <summary>Name of the table. Takes precedence over the index.</summary>
    [Category("Input")]
    [DisplayName("Table Name")]
    [Description("Name of the table. Takes precedence over the index.")]
    public InArgument<string> TableName { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.DeleteTable(
            sheetName,
            new TableRef
            {
                TableName = TableName?.Get(context) ?? string.Empty,
                TableIndex = TableIndex?.Get(context) ?? 0,
            });
}
