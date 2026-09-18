using System.Activities;
using System.ComponentModel;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.PageHeaderFooter;

/// <summary>Writes header or footer text.</summary>
[DisplayName("Insert Header Footer")]
[Description("Writes text into the document's headers or footers, with formatting.")]
public sealed class InsertHeaderFooter : BaseNativeChild
{
    /// <summary>Whether to write headers or footers.</summary>
    [Category("Input")]
    [DisplayName("Insert Type")]
    [Description("Whether to write headers or footers.")]
    public EnumHeadersFooters InsertType { get; set; } = EnumHeadersFooters.Header;

    /// <summary>Text for odd pages.</summary>
    [Category("Input")]
    [DisplayName("Odd Page Text")]
    [Description("Text for odd pages, and for every page when the others are left empty.")]
    public InArgument<string> OddPageText { get; set; } = null!;

    /// <summary>Text for even pages.</summary>
    [Category("Input")]
    [DisplayName("Even Page Text")]
    [Description("Text for even pages.")]
    public InArgument<string> EvenPageText { get; set; } = null!;

    /// <summary>Text for the first page.</summary>
    [Category("Input")]
    [DisplayName("First Page Text")]
    [Description("Text for the first page.")]
    public InArgument<string> FirstPageText { get; set; } = null!;

    /// <summary>Font name.</summary>
    [Category("Input")]
    [DisplayName("Text Font Name")]
    [Description("Font name. Empty leaves the default.")]
    public InArgument<string> TextFontName { get; set; } = null!;

    /// <summary>Font size in points.</summary>
    [Category("Input")]
    [DisplayName("Font Size")]
    [Description("Font size in points. Zero leaves the default.")]
    public InArgument<float> FontSize { get; set; } = null!;

    /// <summary>Bold.</summary>
    [Category("Input")]
    [DisplayName("Font Bold")]
    [Description("Bold.")]
    public EnumBoolean FontBold { get; set; } = EnumBoolean.False;

    /// <summary>Italic.</summary>
    [Category("Input")]
    [DisplayName("Font Italic")]
    [Description("Italic.")]
    public EnumBoolean FontItalic { get; set; } = EnumBoolean.False;

    /// <summary>Underline.</summary>
    [Category("Input")]
    [DisplayName("Font Underline")]
    [Description("Underline.")]
    public EnumBoolean FontUnderline { get; set; } = EnumBoolean.False;

    /// <summary>Strikethrough.</summary>
    [Category("Input")]
    [DisplayName("Font Strikeout")]
    [Description("Strikethrough.")]
    public EnumBoolean FontStrikeout { get; set; } = EnumBoolean.False;

    /// <summary>Horizontal alignment.</summary>
    [Category("Input")]
    [DisplayName("Text Alignment")]
    [Description("Horizontal alignment. Select leaves the default.")]
    public HeaderFooterParagraphAlignment TextAlignment { get; set; } =
        HeaderFooterParagraphAlignment.Select;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.InsertHeaderFooter(InsertType, new HeaderFooterText
        {
            OddPageText = OddPageText?.Get(context) ?? string.Empty,
            EvenPageText = EvenPageText?.Get(context) ?? string.Empty,
            FirstPageText = FirstPageText?.Get(context) ?? string.Empty,
            FontName = TextFontName?.Get(context) ?? string.Empty,
            FontSize = FontSize?.Get(context) ?? 0,
            Bold = FontBold,
            Italic = FontItalic,
            Underline = FontUnderline,
            Strikeout = FontStrikeout,
            Alignment = TextAlignment,
        });
}
