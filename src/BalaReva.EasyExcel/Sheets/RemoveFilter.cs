using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Clears the filter on a table.</summary>
[DisplayName("Remove Filter")]
[Description("Clears the filter on a table.")]
public sealed class RemoveFilter : BaseActivity
{
    /// <summary>Which table on the sheet, numbered from 1.</summary>
    [Category("Input")]
    [DisplayName("Table Index")]
    [Description("Which table on the sheet, numbered from 1.")]
    public InArgument<int> TableIndex { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.RemoveFilter(sheetName, TableIndex?.Get(context) ?? 1);
}
