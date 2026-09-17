using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyImage;

/// <summary>Stacks two images vertically, top to bottom.</summary>
/// <remarks>See the note on <see cref="ImageCombine"/> about the axis choice.</remarks>
[DisplayName("Image Merge")]
[Description("Stacks two images vertically, top to bottom, into one image.")]
public sealed class ImageMerge : BaseWork
{
    /// <summary>Top image.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Path 1")]
    [Description("Full path of the image placed on top.")]
    public InArgument<string> ImagePath1 { get; set; } = null!;

    /// <summary>Bottom image.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Path 2")]
    [Description("Full path of the image placed underneath.")]
    public InArgument<string> ImagePath2 { get; set; } = null!;

    /// <summary>Where to write the result.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Destination Image Path")]
    [Description("Full path to write the merged image to.")]
    public InArgument<string> DesImagePath { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context) =>
        ImageLayout.Join(
            RequireExistingFile(context, ImagePath1, nameof(ImagePath1)),
            RequireExistingFile(context, ImagePath2, nameof(ImagePath2)),
            RequirePath(context, DesImagePath, nameof(DesImagePath)),
            horizontal: false);
}
