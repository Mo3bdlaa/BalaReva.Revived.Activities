using System.Activities;
using System.ComponentModel;
using System.Drawing;

namespace BalaReva.EasyImage;

/// <summary>Crops a rectangle out of an image.</summary>
[DisplayName("Image Cropper")]
[Description("Crops a rectangle out of an image.")]
public sealed class ImageCropper : BaseWork
{
    /// <summary>Image to crop.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Path")]
    [Description("Full path of the image to crop.")]
    public InArgument<string> ImagePath { get; set; } = null!;

    /// <summary>Where to write the result.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Destination Image Path")]
    [Description("Full path to write the cropped image to.")]
    public InArgument<string> DesImagePath { get; set; } = null!;

    /// <summary>Left edge of the crop, in pixels from the left.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("X")]
    [Description("Left edge of the crop rectangle, in pixels from the left of the image.")]
    public InArgument<int> ImgX { get; set; } = null!;

    /// <summary>Top edge of the crop, in pixels from the top.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Y")]
    [Description("Top edge of the crop rectangle, in pixels from the top of the image.")]
    public InArgument<int> ImgY { get; set; } = null!;

    /// <summary>Width of the crop rectangle.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Width")]
    [Description("Width of the crop rectangle, in pixels.")]
    public InArgument<int> ImgWidth { get; set; } = null!;

    /// <summary>Height of the crop rectangle.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Height")]
    [Description("Height of the crop rectangle, in pixels.")]
    public InArgument<int> ImgHeight { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context)
    {
        var source = RequireExistingFile(context, ImagePath, nameof(ImagePath));
        var destination = RequirePath(context, DesImagePath, nameof(DesImagePath));

        using var image = ImageIO.Load(source);
        var crop = new Rectangle(
            ImgX.Get(context), ImgY.Get(context), ImgWidth.Get(context), ImgHeight.Get(context));

        if (crop.Width <= 0 || crop.Height <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(ImgWidth), crop, "The crop width and height must both be greater than zero.");
        }

        // GDI+ would silently clamp a rectangle that runs off the edge, producing a
        // smaller image than asked for. Failing loudly is easier to debug in a robot.
        if (!new Rectangle(0, 0, image.Width, image.Height).Contains(crop))
        {
            throw new ArgumentOutOfRangeException(
                nameof(ImgX), crop,
                $"The crop rectangle falls outside the image, which is {image.Width}x{image.Height}.");
        }

        using var cropped = image.Clone(crop, image.PixelFormat);
        ImageIO.Save(cropped, destination, ImageIO.FormatFromExtension(destination));
    }
}
