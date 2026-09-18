using System.Activities;
using System.ComponentModel;
using System.Drawing;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Fills a range with a colour and a pattern.</summary>
[DisplayName("Fill Color")]
[Description("Fills a range with a colour and a pattern.")]
public sealed class FillColor : BaseActivity
{
    /// <summary>Fill colour.</summary>
    [Category("Input")]
    [DisplayName("Background Color")]
    [Description("Fill colour.")]
    public InArgument<Color> BackgroundColor { get; set; } = null!;

    /// <summary>Range to work over, for example A1:D10.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell Range")]
    [Description("Range to work over, for example A1:D10.")]
    public InArgument<string> CellRange { get; set; } = null!;

    /// <summary>Colour of the pattern drawn over the fill.</summary>
    [Category("Input")]
    [DisplayName("Pattern Color")]
    [Description("Colour of the pattern drawn over the fill.")]
    public InArgument<Color> PatternColor { get; set; } = null!;

    /// <summary>Pattern to draw. None leaves a plain fill.</summary>
    [Category("Input")]
    [DisplayName("Pattern Style")]
    [Description("Pattern to draw. None leaves a plain fill.")]
    public PatternEnum PatternStyle { get; set; } = PatternEnum.None;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.FillColor(
            sheetName,
            Require(context, CellRange, nameof(CellRange)),
            BackgroundColor?.Get(context) ?? Color.Empty,
            PatternColor?.Get(context) ?? Color.Empty,
            PatternStyle);
}
