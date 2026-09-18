using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Clears the filter on a sheet that has no table.</summary>
[DisplayName("Remove Filter Non Table")]
[Description("Clears the filter on a sheet that has no table.")]
public sealed class RemoveFilterNonTable : BaseActivity
{
    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.RemoveFilterNonTable(sheetName);
}
