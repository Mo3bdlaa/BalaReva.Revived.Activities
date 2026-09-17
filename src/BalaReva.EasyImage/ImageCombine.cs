using System.Activities;
using System.ComponentModel;
using System.Drawing;

namespace BalaReva.EasyImage;

/// <summary>Places two images side by side, left to right.</summary>
/// <remarks>
/// The published package ships both an Image Combine and an Image Merge taking the
/// same three arguments, and metadata cannot say which axis each one used. Combine
/// lays the images out horizontally here and Merge stacks them vertically, which at
/// least makes the pair useful and distinct. See docs/REVIVAL.md.
/// </remarks>
[DisplayName("Image Combine")]
[Description("Places two images side by side, left to right, into one image.")]
public sealed class ImageCombine : BaseWork
{
    /// <summary>Left-hand image.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Path 1")]
    [Description("Full path of the image placed on the left.")]
    public InArgument<string> ImagePath1 { get; set; } = null!;

    /// <summary>Right-hand image.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Path 2")]
    [Description("Full path of the image placed on the right.")]
    public InArgument<string> ImagePath2 { get; set; } = null!;

    /// <summary>Where to write the result.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Destination Image Path")]
    [Description("Full path to write the combined image to.")]
    public InArgument<string> DesImagePath { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context) =>
        ImageLayout.Join(
            RequireExistingFile(context, ImagePath1, nameof(ImagePath1)),
            RequireExistingFile(context, ImagePath2, nameof(ImagePath2)),
            RequirePath(context, DesImagePath, nameof(DesImagePath)),
            horizontal: true);
}
