using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Utilities;

namespace BalaReva.Excel.Charts;

/// <summary>Draws a column chart on a sheet.</summary>
/// <remarks>
/// Carries three label options the other chart activities do not, which is how the
/// published package had it.
/// </remarks>
[DisplayName("Column Chart")]
[Description("Draws a column chart from a range of values.")]
public sealed class ColumnChart : BaseChart
{
    /// <summary>Which column chart shape to draw.</summary>
    [Category("Input")]
    [DisplayName("Chart Type")]
    [Description("Which column chart shape to draw.")]
    public ColumnChartEnum ChartType { get; set; } = ColumnChartEnum.ColumnClustered;

    /// <summary>Show the legend key beside each label.</summary>
    [Category("Input")]
    [DisplayName("Show Legend Key")]
    [Description("Show the legend key beside each data label.")]
    public bool ShowLegendKey { get; set; }

    /// <summary>Where value labels sit.</summary>
    [Category("Input")]
    [DisplayName("Show Value Position")]
    [Description("Where value labels sit relative to their columns.")]
    public DataLabelPositionEnum ShowValuePosition { get; set; } = DataLabelPositionEnum.Center;

    /// <summary>Orientation of value labels.</summary>
    [Category("Input")]
    [DisplayName("Show Value Text Orientation")]
    [Description("Orientation of the value labels.")]
    public TextOrientationEnum ShowValueTextOrientation { get; set; } = TextOrientationEnum.Horizontal;

    /// <summary>True when the chart was drawn.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the chart was drawn.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
    {
        var request = Request(context, (int)ChartType);
        request.ShowLegendKey = ShowLegendKey;
        request.ShowValuePosition = ShowValuePosition;
        request.ShowValueTextOrientation = ShowValueTextOrientation;
        workbook.DrawChart(sheetName, request);
    }

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
