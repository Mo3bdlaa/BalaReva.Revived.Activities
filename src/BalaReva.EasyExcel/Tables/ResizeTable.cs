using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Tables;

/// <summary>Resizes a table to a new range.</summary>
[DisplayName("Resize Table")]
[Description("Resizes a table to a new range.")]
public sealed class ResizeTable : BaseActivity
{
    /// <summary>New range for the table, for example A1:D20.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("New Table Range")]
    [Description("New range for the table, for example A1:D20.")]
    public InArgument<string> NewTableRange { get; set; } = null!;

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
        => workbook.ResizeTable(
            sheetName,
            new TableRef
            {
                TableName = TableName?.Get(context) ?? string.Empty,
                TableIndex = TableIndex?.Get(context) ?? 0,
            },
            Require(context, NewTableRange, nameof(NewTableRange)));
}
