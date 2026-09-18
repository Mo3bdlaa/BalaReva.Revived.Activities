using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Deletes every chart on a sheet.</summary>
[DisplayName("Delete All Charts")]
[Description("Deletes every chart on a sheet.")]
public sealed class DeleteAllCharts : BaseActivity
{
    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.DeleteAllCharts(sheetName);
}
