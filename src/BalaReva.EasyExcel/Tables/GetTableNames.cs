using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Tables;

/// <summary>Lists the tables on a sheet.</summary>
[DisplayName("Get Table Names")]
[Description("Lists the tables on a sheet.")]
public sealed class GetTableNames : BaseActivity
{
    /// <summary>The tables, by name.</summary>
    [Category("Output")]
    [DisplayName("Output")]
    [Description("The tables, by name.")]
    public OutArgument<List<string>> Output { get; set; } = null!;

    /// <summary>One row per table, with its name and range.</summary>
    [Category("Output")]
    [DisplayName("Output Table")]
    [Description("One row per table, with its name and range.")]
    public OutArgument<DataTable> OutputTable { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
    {
        var (names, table) = workbook.GetTableNames(sheetName);
        Output.Set(context, names);
        OutputTable.Set(context, table);
    }
}
