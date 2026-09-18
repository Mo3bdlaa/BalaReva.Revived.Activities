namespace BalaReva.Excel.Charts;

/// <summary>Where and how big to place a chart on a sheet, in points.</summary>
public sealed class ChartSize
{
    /// <summary>Distance from the left edge of the sheet.</summary>
    public double Left { get; set; }

    /// <summary>Distance from the top edge of the sheet.</summary>
    public double Top { get; set; }

    /// <summary>Width.</summary>
    public double Width { get; set; }

    /// <summary>Height.</summary>
    public double Height { get; set; }
}

/// <summary>What a chart shows alongside its data.</summary>
public sealed class ShowOptions
{
    /// <summary>Show the series name on each label.</summary>
    public bool ShowSeriesName { get; set; }

    /// <summary>Show the category name on each label.</summary>
    public bool ShowCategoryName { get; set; }

    /// <summary>Show the value on each label.</summary>
    public bool ShowValue { get; set; }

    /// <summary>Show each value as a percentage of the total.</summary>
    public bool ShowPercentage { get; set; }

    /// <summary>Show the bubble size, on chart types that have one.</summary>
    public bool ShowBubbleSize { get; set; }

    /// <summary>Show the chart legend.</summary>
    public bool ShowLegend { get; set; }

    /// <summary>Draw leader lines from labels to their points.</summary>
    public bool HasLeaderLines { get; set; }

    /// <summary>Let Excel choose the label text.</summary>
    public bool AutoText { get; set; }

    /// <summary>Text placed between parts of a label.</summary>
    public string Separator { get; set; } = string.Empty;

    /// <summary>Which labels to show.</summary>
    public DataLabelsEnum DataLabelsType { get; set; } = DataLabelsEnum.ShowNone;

    /// <summary>Whether the legend key appears next to each label.</summary>
    public DataLabelsEnum LegendKey { get; set; } = DataLabelsEnum.ShowNone;
}
