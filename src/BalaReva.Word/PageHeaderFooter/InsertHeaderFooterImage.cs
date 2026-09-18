using System.Activities;
using System.ComponentModel;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.PageHeaderFooter;

/// <summary>Places images in the header or footer.</summary>
[DisplayName("Insert Header Footer Image")]
[Description("Places images in the document's headers or footers.")]
public sealed class InsertHeaderFooterImage : BaseNativeChild
{
    /// <summary>Whether to write headers or footers.</summary>
    [Category("Input")]
    [DisplayName("Insert Type")]
    [Description("Whether to write headers or footers.")]
    public EnumHeadersFooters InsertType { get; set; } = EnumHeadersFooters.Header;

    /// <summary>Image for odd pages.</summary>
    [Category("Input")]
    [DisplayName("Odd Page Image")]
    [Description("Full path of the image for odd pages.")]
    public InArgument<string> OddPageImage { get; set; } = null!;

    /// <summary>Image for even pages.</summary>
    [Category("Input")]
    [DisplayName("Even Page Image")]
    [Description("Full path of the image for even pages.")]
    public InArgument<string> EvenPageImage { get; set; } = null!;

    /// <summary>Image for the first page.</summary>
    [Category("Input")]
    [DisplayName("First Page Image")]
    [Description("Full path of the image for the first page.")]
    public InArgument<string> FirstPageImage { get; set; } = null!;

    /// <summary>Horizontal alignment.</summary>
    [Category("Input")]
    [DisplayName("Image Alignment")]
    [Description("Horizontal alignment. Select leaves the default.")]
    public HeaderFooterParagraphAlignment ImageAlignment { get; set; } =
        HeaderFooterParagraphAlignment.Select;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.InsertHeaderFooterImage(InsertType, new HeaderFooterImages
        {
            OddPageImage = OddPageImage?.Get(context) ?? string.Empty,
            EvenPageImage = EvenPageImage?.Get(context) ?? string.Empty,
            FirstPageImage = FirstPageImage?.Get(context) ?? string.Empty,
            Alignment = ImageAlignment,
        });
}
