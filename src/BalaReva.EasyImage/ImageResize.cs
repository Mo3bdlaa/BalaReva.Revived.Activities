using System.Activities;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace BalaReva.EasyImage;

/// <summary>Resizes an image to an exact width and height.</summary>
[DisplayName("Image Resize")]
[Description("Resizes an image to the given width and height, in pixels.")]
public sealed class ImageResize : BaseWork
{
    /// <summary>Image to resize.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Path")]
    [Description("Full path of the image to resize.")]
    public InArgument<string> ImagePath { get; set; } = null!;

    /// <summary>Where to write the result.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Destination Image Path")]
    [Description("Full path to write the resized image to.")]
    public InArgument<string> DesImagePath { get; set; } = null!;

    /// <summary>Target width in pixels.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Width")]
    [Description("Target width, in pixels.")]
    public InArgument<int> ImageWidth { get; set; } = null!;

    /// <summary>Target height in pixels.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Height")]
    [Description("Target height, in pixels.")]
    public InArgument<int> ImageHeight { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context)
    {
        var source = RequireExistingFile(context, ImagePath, nameof(ImagePath));
        var destination = RequirePath(context, DesImagePath, nameof(DesImagePath));
        var width = ImageWidth.Get(context);
        var height = ImageHeight.Get(context);
        if (width <= 0)
            throw new ArgumentOutOfRangeException(nameof(ImageWidth), width, "Width must be greater than zero.");
        if (height <= 0)
            throw new ArgumentOutOfRangeException(nameof(ImageHeight), height, "Height must be greater than zero.");

        using var image = ImageIO.Load(source);
        using var resized = new Bitmap(width, height);
        using (var graphics = Graphics.FromImage(resized))
        {
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            graphics.DrawImage(image, 0, 0, width, height);
        }

        ImageIO.Save(resized, destination, ImageIO.FormatFromExtension(destination));
    }
}
