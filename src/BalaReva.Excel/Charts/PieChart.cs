using System.Activities;
using System.ComponentModel;

namespace BalaReva.Excel.Charts;

/// <summary>Draws a pie chart on a sheet.</summary>
[DisplayName("PieChart")]
[Description("Draws a pie chart from a range of values.")]
public sealed class PieChart : BaseChart
{
    /// <summary>Which pie chart shape to draw.</summary>
    [Category("Input")]
    [DisplayName("Chart Type")]
    [Description("Which pie chart shape to draw.")]
    public PieChartEnum ChartType { get; set; } = PieChartEnum.Pie;

    /// <summary>True when the chart was drawn.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the chart was drawn.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.DrawChart(sheetName, Request(context, (int)ChartType));

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
