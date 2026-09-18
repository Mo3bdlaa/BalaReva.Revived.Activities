using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.HyperLinks;

/// <summary>Removes a cell's hyperlink.</summary>
[DisplayName("Remove Hyperlink")]
[Description("Removes a cell's hyperlink.")]
public sealed class RemoveHyperlink : BaseActivity
{
    /// <summary>Cell to clear, for example B2.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell")]
    [Description("Cell to clear, for example B2.")]
    public InArgument<string> Cell { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.RemoveHyperlink(
            sheetName, Require(context, Cell, nameof(Cell)));
}
