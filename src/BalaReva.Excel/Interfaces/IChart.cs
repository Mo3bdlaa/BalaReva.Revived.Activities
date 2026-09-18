using System.Activities;
using BalaReva.Excel.Charts;

namespace BalaReva.Excel.Interfaces;

/// <summary>What every chart activity takes.</summary>
public interface IChart
{
    /// <summary>Workbook to draw the chart in.</summary>
    InArgument<string> FilePath { get; set; }

    /// <summary>Sheet to draw the chart on.</summary>
    InArgument<string> SheetName { get; set; }

    /// <summary>Range holding the values to plot.</summary>
    InArgument<string> CellRange { get; set; }

    /// <summary>Title shown above the chart.</summary>
    InArgument<string> ChartTitle { get; set; }

    /// <summary>What the chart shows alongside its data.</summary>
    ShowOptions Options { get; set; }

    /// <summary>Where and how big the chart is.</summary>
    ChartSize Size { get; set; }
}

/// <summary>Position and size of a chart, in points.</summary>
public interface IChartSize
{
    /// <summary>Distance from the left edge of the sheet.</summary>
    double Left { get; set; }

    /// <summary>Distance from the top edge of the sheet.</summary>
    double Top { get; set; }

    /// <summary>Width.</summary>
    double Width { get; set; }

    /// <summary>Height.</summary>
    double Height { get; set; }
}

/// <summary>What a chart shows alongside its data.</summary>
public interface IShowOptions
{
    /// <summary>Show the series name on each label.</summary>
    bool ShowSeriesName { get; set; }

    /// <summary>Show the category name on each label.</summary>
    bool ShowCategoryName { get; set; }

    /// <summary>Show the value on each label.</summary>
    bool ShowValue { get; set; }

    /// <summary>Show each value as a percentage of the total.</summary>
    bool ShowPercentage { get; set; }

    /// <summary>Show the bubble size, on chart types that have one.</summary>
    bool ShowBubbleSize { get; set; }

    /// <summary>Show the chart legend.</summary>
    bool ShowLegend { get; set; }

    /// <summary>Draw leader lines from labels to their points.</summary>
    bool HasLeaderLines { get; set; }

    /// <summary>Let Excel choose the label text.</summary>
    bool AutoText { get; set; }

    /// <summary>Text placed between parts of a label.</summary>
    string Separator { get; set; }

    /// <summary>Which labels to show.</summary>
    DataLabelsEnum DataLabelsType { get; set; }

    /// <summary>Whether the legend key appears next to each label.</summary>
    DataLabelsEnum LegendKey { get; set; }
}

/// <summary>A bar chart's shape.</summary>
public interface IBarChart
{
    /// <summary>Which bar chart shape to draw.</summary>
    BarChartEnum ChartType { get; set; }
}

/// <summary>A column chart's shape.</summary>
public interface IColumnChart
{
    /// <summary>Which column chart shape to draw.</summary>
    ColumnChartEnum ChartType { get; set; }
}

/// <summary>A line chart's shape.</summary>
public interface ILineChart
{
    /// <summary>Which line chart shape to draw.</summary>
    LineChartEnum ChartType { get; set; }
}

/// <summary>A pie chart's shape.</summary>
public interface IPieChart
{
    /// <summary>Which pie chart shape to draw.</summary>
    PieChartEnum ChartType { get; set; }
}
