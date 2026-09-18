using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Selects a range.</summary>
[DisplayName("Select Cell")]
[Description("Selects a range.")]
public sealed class SelectCell : BaseActivity
{
    /// <summary>Range to select, for example A1:D10.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell Range")]
    [Description("Range to select, for example A1:D10.")]
    public InArgument<string> CellRange { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.SelectCell(
            sheetName, Require(context, CellRange, nameof(CellRange)));
}
