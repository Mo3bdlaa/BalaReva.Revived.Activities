using System.Activities;
using System.ComponentModel;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.ImageDocument;

/// <summary>Saves the document's inline images to a folder.</summary>
[DisplayName("Image Extract")]
[Description("Saves the document's inline images into a folder.")]
public sealed class ImageExtract : BaseNativeChild
{
    /// <summary>Folder to save the images into.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Folder")]
    [Description("Folder to save the images into. Created if it does not exist.")]
    public InArgument<string> ImageFolder { get; set; } = null!;

    /// <summary>Image format to save in.</summary>
    [Category("Input")]
    [DisplayName("File Extension")]
    [Description("Format to save the images in.")]
    public EnumFileExtension FileExtension { get; set; } = EnumFileExtension.Png;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.ExtractImages(Require(context, ImageFolder, nameof(ImageFolder)), FileExtension);
}
