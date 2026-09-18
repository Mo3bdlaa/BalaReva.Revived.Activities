using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint.Utilities;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Saves a slide images into a folder.</summary>
[DisplayName("Image Extractor")]
[Description("Saves a slide images into a folder.")]
public sealed class ImageExtractor : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>Folder to save the images into. Created if it does not exist.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Directory")]
    [Description("Folder to save the images into. Created if it does not exist.")]
    public InArgument<string> ImageDirectory { get; set; } = null!;

    /// <summary>Format to save the images in.</summary>
    [Category("Input")]
    [DisplayName("Image File Format")]
    [Description("Format to save the images in.")]
    public ImageFileFormatEnum ImageFileFormat { get; set; } = ImageFileFormatEnum.PNG;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.ImageExtractor(
            SlideIndex.Get(context),
            Require(context, ImageDirectory, nameof(ImageDirectory)),
            ImageFileFormat);
}
