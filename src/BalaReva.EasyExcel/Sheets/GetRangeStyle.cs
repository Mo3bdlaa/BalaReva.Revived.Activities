using System.Activities;
using System.ComponentModel;
using System.Drawing;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Reads the font and fill of a range's first cell.</summary>
/// <remarks>
/// Font_Style, FontScript and FontUnderLine are plain properties here rather than output
/// arguments, which is how the published package declared them, so they cannot report
/// what was read. The colours, name and size can.
/// </remarks>
[DisplayName("Get Range Style")]
[Description("Reads the font and fill of a range's first cell.")]
public sealed class GetRangeStyle : BaseActivity
{
    /// <summary>Fill colour. Empty leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Background Color")]
    [Description("Fill colour. Empty leaves it alone.")]
    public OutArgument<Color> BackgroundColor { get; set; } = null!;

    /// <summary>Range to work over, for example A1:D10.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell Range")]
    [Description("Range to work over, for example A1:D10.")]
    public InArgument<string> CellRange { get; set; } = null!;

    /// <summary>Bold, italic, both, or neither.</summary>
    [Category("Input")]
    [DisplayName("Font Style")]
    [Description("Bold, italic, both, or neither.")]
    public FontStyleEnum Font_Style { get; set; } = FontStyleEnum.None;

    /// <summary>Font colour. Empty leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Font Color")]
    [Description("Font colour. Empty leaves it alone.")]
    public OutArgument<Color> FontColor { get; set; } = null!;

    /// <summary>Font name. Empty leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Font Name")]
    [Description("Font name. Empty leaves it alone.")]
    public OutArgument<string> FontName { get; set; } = null!;

    /// <summary>Superscript or subscript.</summary>
    [Category("Input")]
    [DisplayName("Font Script")]
    [Description("Superscript or subscript.")]
    public FontScriptEnum FontScript { get; set; } = FontScriptEnum.None;

    /// <summary>Font size in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Font Size")]
    [Description("Font size in points. Zero leaves it alone.")]
    public OutArgument<double> FontSize { get; set; } = null!;

    /// <summary>Underline style.</summary>
    [Category("Input")]
    [DisplayName("Font Under Line")]
    [Description("Underline style.")]
    public FontUnderLineEnum FontUnderLine { get; set; } = FontUnderLineEnum.None;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
    {
        var style = workbook.GetRangeStyle(
            sheetName, Require(context, CellRange, nameof(CellRange)));

        FontName.Set(context, style.FontName);
        FontSize.Set(context, style.FontSize);
        FontColor.Set(context, style.FontColor);
        BackgroundColor.Set(context, style.BackgroundColor);
    }
}
