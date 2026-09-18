using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Tables;

/// <summary>Reports whether a table is on a sheet.</summary>
[DisplayName("Table Exists")]
[Description("Reports whether a table is on a sheet.")]
public sealed class TableExists : BaseActivity
{
    /// <summary>True when a table of that name is on the sheet.</summary>
    [Category("Output")]
    [DisplayName("Result")]
    [Description("True when a table of that name is on the sheet.")]
    public OutArgument<bool> Result { get; set; } = null!;

    /// <summary>Name of the table to look for.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Table Name")]
    [Description("Name of the table to look for.")]
    public InArgument<string> TableName { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => Result.Set(context, workbook.TableExists(
            sheetName, Require(context, TableName, nameof(TableName))));
}
