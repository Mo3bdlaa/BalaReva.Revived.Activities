using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyImage;

/// <summary>Re-encodes an image as JPEG at a chosen quality.</summary>
[DisplayName("Image Compression")]
[Description("Re-encodes an image as JPEG at the given quality, from 0 to 100.")]
public sealed class ImageCompression : BaseWork
{
    /// <summary>Image to compress.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Path")]
    [Description("Full path of the image to compress.")]
    public InArgument<string> ImagePath { get; set; } = null!;

    /// <summary>Where to write the result.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Destination Image Path")]
    [Description("Full path to write the compressed image to.")]
    public InArgument<string> DesImagePath { get; set; } = null!;

    /// <summary>JPEG quality, 0 to 100. Higher is better looking and larger.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Quality")]
    [Description("JPEG quality from 0 to 100. Higher means better quality and a larger file.")]
    public InArgument<int> ImageQuality { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context)
    {
        var source = RequireExistingFile(context, ImagePath, nameof(ImagePath));
        var destination = RequirePath(context, DesImagePath, nameof(DesImagePath));

        using var image = ImageIO.Load(source);
        ImageIO.SaveJpeg(image, destination, ImageQuality.Get(context));
    }
}
