using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Charts;

/// <summary>Deletes every chart on a sheet.</summary>
[DisplayName("Chart Delete All")]
[Description("Deletes every chart on a sheet.")]
public sealed class ChartDeleteAll : BaseActivity
{
    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.ChartDeleteAll(sheetName);
}
