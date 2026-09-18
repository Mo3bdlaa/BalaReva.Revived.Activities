using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Sets up a sheet for printing.</summary>
/// <remarks>
/// Twenty-one arguments, and every one of them optional: an empty string, a zero and a
/// None enum each leave that setting as the sheet already had it.
/// </remarks>
[DisplayName("Page Setup")]
[Description("Sets up a sheet for printing.")]
public sealed class PageSetup : BaseActivity
{
    /// <summary>Bottom margin in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Bottom Margin")]
    [Description("Bottom margin in points. Zero leaves it alone.")]
    public InArgument<int> BottomMargin { get; set; } = null!;

    /// <summary>Footer text, centre.</summary>
    [Category("Input")]
    [DisplayName("Center Footer")]
    [Description("Footer text, centre.")]
    public InArgument<string> CenterFooter { get; set; } = null!;

    /// <summary>Header text, centre.</summary>
    [Category("Input")]
    [DisplayName("Center Header")]
    [Description("Header text, centre.")]
    public InArgument<string> CenterHeader { get; set; } = null!;

    /// <summary>Centre the printout on the page.</summary>
    [Category("Input")]
    [DisplayName("Center On Page")]
    [Description("Centre the printout on the page.")]
    public CenterOnPageEnum CenterOnPage { get; set; } = CenterOnPageEnum.None;

    /// <summary>Whether to fit to a page height at all.</summary>
    [Category("Input")]
    [DisplayName("Fit To Pages Tall")]
    [Description("Whether to fit to a page height at all.")]
    public TrueFaleNoneEnum FitToPagesTall { get; set; } = TrueFaleNoneEnum.None;

    /// <summary>How many pages wide to fit to. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Fit To Pages Wide")]
    [Description("How many pages wide to fit to. Zero leaves it alone.")]
    public InArgument<int> FitToPagesWide { get; set; } = null!;

    /// <summary>Footer margin in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Footer Margin")]
    [Description("Footer margin in points. Zero leaves it alone.")]
    public InArgument<int> FooterMargin { get; set; } = null!;

    /// <summary>Header margin in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Header Margin")]
    [Description("Header margin in points. Zero leaves it alone.")]
    public InArgument<int> HeaderMargin { get; set; } = null!;

    /// <summary>Footer text, left.</summary>
    [Category("Input")]
    [DisplayName("Left Footer")]
    [Description("Footer text, left.")]
    public InArgument<string> LeftFooter { get; set; } = null!;

    /// <summary>Header text, left.</summary>
    [Category("Input")]
    [DisplayName("Left Header")]
    [Description("Header text, left.")]
    public InArgument<string> LeftHeader { get; set; } = null!;

    /// <summary>Left margin in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Left Margin")]
    [Description("Left margin in points. Zero leaves it alone.")]
    public InArgument<int> LeftMargin { get; set; } = null!;

    /// <summary>Down then over, or over then down.</summary>
    [Category("Input")]
    [DisplayName("Page Order")]
    [Description("Down then over, or over then down.")]
    public PageOrderEnum PageOrder { get; set; } = PageOrderEnum.None;

    /// <summary>Portrait or landscape.</summary>
    [Category("Input")]
    [DisplayName("Page Orientation")]
    [Description("Portrait or landscape.")]
    public PageOrientationEnum PageOrientation { get; set; } = PageOrientationEnum.None;

    /// <summary>Paper size.</summary>
    [Category("Input")]
    [DisplayName("Page Size")]
    [Description("Paper size.")]
    public PaperSizeEnum PageSize { get; set; } = PaperSizeEnum.PaperNone;

    /// <summary>Print the gridlines.</summary>
    [Category("Input")]
    [DisplayName("Print Gridlines")]
    [Description("Print the gridlines.")]
    public TrueFaleNoneEnum PrintGridlines { get; set; } = TrueFaleNoneEnum.None;

    /// <summary>Scale the printout rather than fitting it to pages.</summary>
    [Category("Input")]
    [DisplayName("Print Zoom")]
    [Description("Scale the printout rather than fitting it to pages.")]
    public TrueFaleNoneEnum PrintZoom { get; set; } = TrueFaleNoneEnum.None;

    /// <summary>Footer text, right.</summary>
    [Category("Input")]
    [DisplayName("Right Footer")]
    [Description("Footer text, right.")]
    public InArgument<string> RightFooter { get; set; } = null!;

    /// <summary>Header text, right.</summary>
    [Category("Input")]
    [DisplayName("Right Header")]
    [Description("Header text, right.")]
    public InArgument<string> RightHeader { get; set; } = null!;

    /// <summary>Right margin in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Right Margin")]
    [Description("Right margin in points. Zero leaves it alone.")]
    public InArgument<int> RightMargin { get; set; } = null!;

    /// <summary>Top margin in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Top Margin")]
    [Description("Top margin in points. Zero leaves it alone.")]
    public InArgument<int> TopMargin { get; set; } = null!;

    /// <summary>Zoom percentage. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Zoom Level")]
    [Description("Zoom percentage. Zero leaves it alone.")]
    public InArgument<int> ZoomLevel { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.PageSetup(sheetName, new PageSetupRequest
        {
            LeftHeader = LeftHeader?.Get(context) ?? string.Empty,
            CenterHeader = CenterHeader?.Get(context) ?? string.Empty,
            RightHeader = RightHeader?.Get(context) ?? string.Empty,
            LeftFooter = LeftFooter?.Get(context) ?? string.Empty,
            CenterFooter = CenterFooter?.Get(context) ?? string.Empty,
            RightFooter = RightFooter?.Get(context) ?? string.Empty,
            LeftMargin = LeftMargin?.Get(context) ?? 0,
            RightMargin = RightMargin?.Get(context) ?? 0,
            TopMargin = TopMargin?.Get(context) ?? 0,
            BottomMargin = BottomMargin?.Get(context) ?? 0,
            HeaderMargin = HeaderMargin?.Get(context) ?? 0,
            FooterMargin = FooterMargin?.Get(context) ?? 0,
            CenterOnPage = CenterOnPage,
            PageOrientation = PageOrientation,
            PageOrder = PageOrder,
            PageSize = PageSize,
            FitToPagesWide = FitToPagesWide?.Get(context) ?? 0,
            FitToPagesTall = FitToPagesTall,
            PrintGridlines = PrintGridlines,
            PrintZoom = PrintZoom,
            ZoomLevel = ZoomLevel?.Get(context) ?? 0,
        });
}
