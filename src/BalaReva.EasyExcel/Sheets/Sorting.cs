using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Sorts a table by the chosen columns.</summary>
[DisplayName("Sorting")]
[Description("Sorts a table by the chosen columns.")]
public sealed class Sorting : BaseActivity
{
    /// <summary>Columns to sort by, by position, numbered from 1.</summary>
    [Category("Input")]
    [DisplayName("Column Indexes")]
    [Description("Columns to sort by, by position, numbered from 1.")]
    public InArgument<int[]> ColumnIndexes { get; set; } = null!;

    /// <summary>Columns to sort by, by name.</summary>
    [Category("Input")]
    [DisplayName("Column Names")]
    [Description("Columns to sort by, by name.")]
    public InArgument<string[]> ColumnNames { get; set; } = null!;

    /// <summary>Ascending or descending.</summary>
    [Category("Input")]
    [DisplayName("Sorting Order")]
    [Description("Ascending or descending.")]
    public SortingTableEnum SortingOrder { get; set; } = SortingTableEnum.Ascending;

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
        => workbook.Sorting(
            sheetName,
            new TableRef
            {
                TableName = TableName?.Get(context) ?? string.Empty,
                TableIndex = TableIndex?.Get(context) ?? 0,
            },
            ColumnNames?.Get(context) ?? [],
            ColumnIndexes?.Get(context) ?? [],
            SortingOrder);
}
