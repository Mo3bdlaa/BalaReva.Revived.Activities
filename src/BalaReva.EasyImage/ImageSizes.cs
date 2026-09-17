using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyImage;

/// <summary>Reads the pixel dimensions of an image.</summary>
[DisplayName("Image Sizes")]
[Description("Reads the width and height of an image, in pixels.")]
public sealed class ImageSizes : BaseWork
{
    /// <summary>Image to measure.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Path")]
    [Description("Full path of the image to measure.")]
    public InArgument<string> ImagePath { get; set; } = null!;

    /// <summary>Width in pixels.</summary>
    [Category("Output")]
    [DisplayName("Image Width")]
    [Description("Width of the image, in pixels.")]
    public OutArgument<int> ImageWidth { get; set; } = null!;

    /// <summary>Height in pixels.</summary>
    [Category("Output")]
    [DisplayName("Image Height")]
    [Description("Height of the image, in pixels.")]
    public OutArgument<int> ImageHeight { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context)
    {
        using var image = ImageIO.Load(RequireExistingFile(context, ImagePath, nameof(ImagePath)));
        ImageWidth.Set(context, image.Width);
        ImageHeight.Set(context, image.Height);
    }
}
