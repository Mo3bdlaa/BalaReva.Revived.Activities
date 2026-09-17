using System.Activities;
using System.ComponentModel;
using System.Drawing;

namespace BalaReva.EasyImage;

/// <summary>Rotates or flips an image.</summary>
[DisplayName("Image Rotate")]
[Description("Rotates or flips an image by a quarter turn.")]
public sealed class ImageRotate : BaseWork
{
    /// <summary>Image to rotate.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Path")]
    [Description("Full path of the image to rotate.")]
    public InArgument<string> ImagePath { get; set; } = null!;

    /// <summary>Where to write the result.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Destination Image Path")]
    [Description("Full path to write the rotated image to.")]
    public InArgument<string> DesImagePath { get; set; } = null!;

    /// <summary>Rotation and flip to apply.</summary>
    [Category("Input")]
    [DisplayName("Flip Type")]
    [Description("The rotation and flip to apply.")]
    public RotateFlipType FlipType { get; set; } = RotateFlipType.RotateNoneFlipNone;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context)
    {
        var source = RequireExistingFile(context, ImagePath, nameof(ImagePath));
        var destination = RequirePath(context, DesImagePath, nameof(DesImagePath));

        using var image = ImageIO.Load(source);
        image.RotateFlip(FlipType);
        ImageIO.Save(image, destination, ImageIO.FormatFromExtension(destination));
    }
}
