using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.Charts;

/// <summary>Shared arguments for every chart activity.</summary>
public abstract class BaseChart : ExcelCore
{
    /// <summary>Range holding the values to plot.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell Range")]
    [Description("Range holding the values to plot, for example A1:B10.")]
    public InArgument<string> CellRange { get; set; } = null!;

    /// <summary>Range holding the category labels.</summary>
    [Category("Input")]
    [DisplayName("Legend Range")]
    [Description("Range holding the category labels along the axis.")]
    public InArgument<string> LegendRange { get; set; } = null!;

    /// <summary>Title shown above the chart.</summary>
    [Category("Input")]
    [DisplayName("Chart Title")]
    [Description("Title shown above the chart. Empty shows none.")]
    public InArgument<string> ChartTitle { get; set; } = null!;

    /// <summary>Path to save a picture of the chart to.</summary>
    [Category("Input")]
    [DisplayName("Image Copy")]
    [Description("Full path to save a picture of the chart to. Empty saves none.")]
    public InArgument<string> ImageCopy { get; set; } = null!;

    /// <summary>What the chart shows alongside its data.</summary>
    [Category("Input")]
    [DisplayName("Options")]
    [Description("What the chart shows alongside its data.")]
    public ShowOptions Options { get; set; } = new();

    /// <summary>Where and how big the chart is.</summary>
    [Category("Input")]
    [DisplayName("Size")]
    [Description("Where the chart sits on the sheet and how big it is, in points.")]
    public ChartSize Size { get; set; } = new();

    /// <summary>Gathers the shared arguments into a request for the service.</summary>
    protected ChartRequest Request(CodeActivityContext context, int chartType) => new()
    {
        ChartType = chartType,
        CellRange = Require(context, CellRange, nameof(CellRange)),
        LegendRange = LegendRange?.Get(context) ?? string.Empty,
        ChartTitle = ChartTitle?.Get(context) ?? string.Empty,
        ImageCopy = ImageCopy?.Get(context) ?? string.Empty,
        Options = Options ?? new ShowOptions(),
        Size = Size ?? new ChartSize(),
    };
}
