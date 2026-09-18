using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Outline;

/// <summary>Collapses every outline group on a sheet.</summary>
[DisplayName("Collapse All Group")]
[Description("Collapses every outline group on a sheet.")]
public sealed class CollapseAllGroup : BaseActivity
{
    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.CollapseAllGroup(sheetName);
}
