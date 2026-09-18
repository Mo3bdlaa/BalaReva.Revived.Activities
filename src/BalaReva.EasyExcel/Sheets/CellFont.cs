using System.Activities;
using System.ComponentModel;
using System.Drawing;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Applies font and fill to a range.</summary>
/// <remarks>
/// Every setting is optional: an empty name, a zero size, an empty colour and a None enum
/// each leave that setting as the range already had it.
/// </remarks>
[DisplayName("Cell Font")]
[Description("Applies font and fill to a range.")]
public sealed class CellFont : BaseActivity
{
    /// <summary>Fill colour. Empty leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Background Color")]
    [Description("Fill colour. Empty leaves it alone.")]
    public InArgument<Color> BackgroundColor { get; set; } = null!;

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
    public InArgument<Color> FontColor { get; set; } = null!;

    /// <summary>Font name. Empty leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Font Name")]
    [Description("Font name. Empty leaves it alone.")]
    public InArgument<string> FontName { get; set; } = null!;

    /// <summary>Superscript or subscript.</summary>
    [Category("Input")]
    [DisplayName("Font Script")]
    [Description("Superscript or subscript.")]
    public FontScriptEnum FontScript { get; set; } = FontScriptEnum.None;

    /// <summary>Font size in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Font Size")]
    [Description("Font size in points. Zero leaves it alone.")]
    public InArgument<double> FontSize { get; set; } = null!;

    /// <summary>Underline style.</summary>
    [Category("Input")]
    [DisplayName("Font Under Line")]
    [Description("Underline style.")]
    public FontUnderLineEnum FontUnderLine { get; set; } = FontUnderLineEnum.None;

    /// <summary>Strikethrough.</summary>
    [Category("Input")]
    [DisplayName("Strikethrough")]
    [Description("Strikethrough.")]
    public SelectionYesNoNone Strikethrough { get; set; } = SelectionYesNoNone.None;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.CellFont(
            sheetName,
            Require(context, CellRange, nameof(CellRange)),
            new CellFontRequest
            {
                FontName = FontName?.Get(context) ?? string.Empty,
                FontSize = FontSize?.Get(context) ?? 0,
                FontStyle = Font_Style,
                FontUnderLine = FontUnderLine,
                FontScript = FontScript,
                Strikethrough = Strikethrough,
                FontColor = FontColor?.Get(context) ?? Color.Empty,
                BackgroundColor = BackgroundColor?.Get(context) ?? Color.Empty,
            });
}
