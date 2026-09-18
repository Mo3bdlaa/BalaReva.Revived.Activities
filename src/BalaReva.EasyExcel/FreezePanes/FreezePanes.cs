using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.FreezePanes;

/// <summary>Freezes or unfreezes the panes at a cell.</summary>
[DisplayName("Freeze Panes")]
[Description("Freezes or unfreezes the panes at a cell.")]
public sealed class FreezePanes : BaseActivity
{
    /// <summary>Cell to freeze at, for example B2.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell Range")]
    [Description("Cell to freeze at, for example B2.")]
    public InArgument<string> CellRange { get; set; } = null!;

    /// <summary>Whether to freeze or to unfreeze.</summary>
    [Category("Input")]
    [DisplayName("Freeze Option")]
    [Description("Whether to freeze or to unfreeze.")]
    public FreezePanesEnum FreezeOption { get; set; } = FreezePanesEnum.Freeze;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.FreezePanes(
            sheetName, Require(context, CellRange, nameof(CellRange)), FreezeOption);
}
