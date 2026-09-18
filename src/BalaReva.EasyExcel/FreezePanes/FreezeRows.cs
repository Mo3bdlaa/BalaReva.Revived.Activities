using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.FreezePanes;

/// <summary>Freezes or unfreezes a sheet's first rows.</summary>
[DisplayName("Freeze Rows")]
[Description("Freezes or unfreezes a sheet's first rows.")]
public sealed class FreezeRows : BaseActivity
{
    /// <summary>Whether to freeze or to unfreeze.</summary>
    [Category("Input")]
    [DisplayName("Freeze Option")]
    [Description("Whether to freeze or to unfreeze.")]
    public FreezePanesEnum FreezeOption { get; set; } = FreezePanesEnum.Freeze;

    /// <summary>How many rows to freeze.</summary>
    [Category("Input")]
    [DisplayName("No Rows")]
    [Description("How many rows to freeze.")]
    public InArgument<int> NoRows { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.FreezeRows(sheetName, NoRows?.Get(context) ?? 1, FreezeOption);
}
