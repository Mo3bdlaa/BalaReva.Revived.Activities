using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.FreezePanes;

/// <summary>Freezes or unfreezes a sheet's first columns.</summary>
[DisplayName("Freeze Columns")]
[Description("Freezes or unfreezes a sheet's first columns.")]
public sealed class FreezeColumns : BaseActivity
{
    /// <summary>Whether to freeze or to unfreeze.</summary>
    [Category("Input")]
    [DisplayName("Freeze Option")]
    [Description("Whether to freeze or to unfreeze.")]
    public FreezePanesEnum FreezeOption { get; set; } = FreezePanesEnum.Freeze;

    /// <summary>How many columns to freeze.</summary>
    [Category("Input")]
    [DisplayName("No Columns")]
    [Description("How many columns to freeze.")]
    public InArgument<int> NoColumns { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.FreezeColumns(sheetName, NoColumns?.Get(context) ?? 1, FreezeOption);
}
