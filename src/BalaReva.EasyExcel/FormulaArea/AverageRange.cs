using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.FormulaArea;

/// <summary>Averages the numbers in a range.</summary>
[DisplayName("Average Range")]
[Description("Averages the numbers in a range.")]
public sealed class AverageRange : BaseActivity
{
    /// <summary>Range to work over, for example A1:D10.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell Range")]
    [Description("Range to work over, for example A1:D10.")]
    public InArgument<string> CellRange { get; set; } = null!;

    /// <summary>The result.</summary>
    [Category("Output")]
    [DisplayName("Output")]
    [Description("The result.")]
    public OutArgument<double> Output { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => Output.Set(context, workbook.RangeFunction(
            sheetName,
            Require(context, CellRange, nameof(CellRange)),
            RangeFunctionKind.Average));
}
