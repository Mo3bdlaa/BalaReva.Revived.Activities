using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Refreshes every data connection in the workbook.</summary>
[DisplayName("Refresh All")]
[Description("Refreshes every data connection in the workbook.")]
public sealed class RefreshAll : BaseActivity
{
    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.RefreshAll();
}
