using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.Outline;

/// <summary>Ungroups a range by row or by column.</summary>
[DisplayName("Un Group")]
[Description("Ungroups a range by row or by column.")]
public sealed class UnGroup : BaseActivity
{
    /// <summary>Range to work over, for example A1:D10.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell Range")]
    [Description("Range to work over, for example A1:D10.")]
    public InArgument<string> CellRange { get; set; } = null!;

    /// <summary>Whether to work by row or by column.</summary>
    [Category("Input")]
    [DisplayName("Group Type")]
    [Description("Whether to work by row or by column.")]
    public GroupEnum GroupType { get; set; } = GroupEnum.Rows;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.UnGroup(
            sheetName, Require(context, CellRange, nameof(CellRange)), GroupType);
}
