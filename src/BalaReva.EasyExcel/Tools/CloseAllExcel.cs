using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Tools;

/// <summary>Closes every running Excel instance, saving nothing.</summary>
/// <remarks>
/// Stands outside any scope, and will close the workbook a scope is holding along with
/// everything else.
/// </remarks>
[DisplayName("Close All Excel")]
[Description("Closes every running Excel instance, saving nothing.")]
public sealed class CloseAllExcel : CodeActivity
{
    /// <inheritdoc />
    protected override void Execute(CodeActivityContext context)
        => (context.GetExtension<IExcelService>() ?? ExcelService.Instance).CloseAllExcel();
}
