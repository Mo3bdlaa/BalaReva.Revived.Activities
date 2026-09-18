using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Charts;

/// <summary>Moves and resizes a chart.</summary>
/// <remarks>
/// Every measurement is in points, and zero leaves that one alone.
/// </remarks>
[DisplayName("Chart Format")]
[Description("Moves and resizes a chart.")]
public sealed class ChartFormat : BaseActivity
{
    /// <summary>Height in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Chart Height")]
    [Description("Height in points. Zero leaves it alone.")]
    public InArgument<double> ChartHeight { get; set; } = null!;

    /// <summary>Position of the chart on the sheet, numbered from 1.</summary>
    [Category("Input")]
    [DisplayName("Chart Index")]
    [Description("Position of the chart on the sheet, numbered from 1.")]
    public InArgument<int> ChartIndex { get; set; } = null!;

    /// <summary>Distance from the left edge of the sheet, in points.</summary>
    [Category("Input")]
    [DisplayName("Chart Left")]
    [Description("Distance from the left edge of the sheet, in points.")]
    public InArgument<double> ChartLeft { get; set; } = null!;

    /// <summary>Name of the chart. Takes precedence over the index.</summary>
    [Category("Input")]
    [DisplayName("Chart Name")]
    [Description("Name of the chart. Takes precedence over the index.")]
    public InArgument<string> ChartName { get; set; } = null!;

    /// <summary>Distance from the top edge of the sheet, in points.</summary>
    [Category("Input")]
    [DisplayName("Chart Top")]
    [Description("Distance from the top edge of the sheet, in points.")]
    public InArgument<double> ChartTop { get; set; } = null!;

    /// <summary>Width in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Chart Width")]
    [Description("Width in points. Zero leaves it alone.")]
    public InArgument<double> ChartWidth { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.ChartFormat(
            new ChartRef
            {
                SheetName = sheetName,
                ChartName = ChartName?.Get(context) ?? string.Empty,
                ChartIndex = ChartIndex?.Get(context) ?? 0,
            },
            new ChartBounds
            {
                Left = ChartLeft?.Get(context) ?? 0,
                Top = ChartTop?.Get(context) ?? 0,
                Width = ChartWidth?.Get(context) ?? 0,
                Height = ChartHeight?.Get(context) ?? 0,
            });
}
