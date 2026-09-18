using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Reports whether a cell is part of a merged cell.</summary>
[DisplayName("Is Merged Cell")]
[Description("Reports whether a cell is part of a merged cell.")]
public sealed class IsMergedCell : BaseActivity
{
    /// <summary>Cell to test, for example B2.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell")]
    [Description("Cell to test, for example B2.")]
    public InArgument<string> Cell { get; set; } = null!;

    /// <summary>True when the cell is part of a merged cell.</summary>
    [Category("Output")]
    [DisplayName("Result")]
    [Description("True when the cell is part of a merged cell.")]
    public OutArgument<bool> Result { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => Result.Set(context, workbook.IsMergedCell(
            sheetName, Require(context, Cell, nameof(Cell))));
}
