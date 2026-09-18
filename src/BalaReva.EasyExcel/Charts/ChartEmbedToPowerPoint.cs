using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Charts;

/// <summary>Copies a chart onto a slide of a presentation.</summary>
[DisplayName("Chart Embed To PowerPoint")]
[Description("Copies a chart onto a slide of a presentation.")]
public sealed class ChartEmbedToPowerPoint : BaseActivity
{
    /// <summary>Height in points. Zero keeps the chart as it is.</summary>
    [Category("Input")]
    [DisplayName("Chart Height")]
    [Description("Height in points. Zero keeps the chart as it is.")]
    public InArgument<float> ChartHeight { get; set; } = null!;

    /// <summary>Position of the chart on the sheet, numbered from 1.</summary>
    [Category("Input")]
    [DisplayName("Chart Index")]
    [Description("Position of the chart on the sheet, numbered from 1.")]
    public InArgument<int> ChartIndex { get; set; } = null!;

    /// <summary>Distance from the left edge of the slide, in points.</summary>
    [Category("Input")]
    [DisplayName("Chart Left")]
    [Description("Distance from the left edge of the slide, in points.")]
    public InArgument<float> ChartLeft { get; set; } = null!;

    /// <summary>Name of the chart. Takes precedence over the index.</summary>
    [Category("Input")]
    [DisplayName("Chart Name")]
    [Description("Name of the chart. Takes precedence over the index.")]
    public InArgument<string> ChartName { get; set; } = null!;

    /// <summary>Distance from the top edge of the slide, in points.</summary>
    [Category("Input")]
    [DisplayName("Chart Top")]
    [Description("Distance from the top edge of the slide, in points.")]
    public InArgument<float> ChartTop { get; set; } = null!;

    /// <summary>Width in points. Zero keeps the chart as it is.</summary>
    [Category("Input")]
    [DisplayName("Chart Width")]
    [Description("Width in points. Zero keeps the chart as it is.")]
    public InArgument<float> ChartWidth { get; set; } = null!;

    /// <summary>Full path of the presentation to paste into.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Ppt File")]
    [Description("Full path of the presentation to paste into.")]
    public InArgument<string> PptFile { get; set; } = null!;

    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.ChartEmbedToPowerPoint(
            new ChartRef
            {
                SheetName = sheetName,
                ChartName = ChartName?.Get(context) ?? string.Empty,
                ChartIndex = ChartIndex?.Get(context) ?? 0,
            },
            Require(context, PptFile, nameof(PptFile)),
            SlideIndex?.Get(context) ?? 1,
            new ChartBounds
            {
                Left = ChartLeft?.Get(context) ?? 0,
                Top = ChartTop?.Get(context) ?? 0,
                Width = ChartWidth?.Get(context) ?? 0,
                Height = ChartHeight?.Get(context) ?? 0,
            });
}
