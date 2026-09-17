using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyImage;

/// <summary>Converts an image to another file format.</summary>
[DisplayName("Image Converter")]
[Description("Converts an image to another file format.")]
public sealed class ImageConverter : BaseWork
{
    /// <summary>Image to convert.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Path")]
    [Description("Full path of the image to convert.")]
    public InArgument<string> ImagePath { get; set; } = null!;

    /// <summary>Where to write the result.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Destination Image Path")]
    [Description("Full path to write the converted image to.")]
    public InArgument<string> DesImagePath { get; set; } = null!;

    /// <summary>Format to convert to.</summary>
    /// <remarks>
    /// This wins over the destination file's extension, so converting to
    /// <see cref="EnumFileFormat.PNG"/> with a .jpg destination really does write PNG
    /// bytes. That mismatch is worth avoiding in a workflow, but silently ignoring
    /// the property the user set would be worse.
    /// </remarks>
    [Category("Input")]
    [DisplayName("File Format")]
    [Description("Format to convert to. Takes precedence over the destination file extension.")]
    public EnumFileFormat FileFormat { get; set; } = EnumFileFormat.PNG;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context)
    {
        var source = RequireExistingFile(context, ImagePath, nameof(ImagePath));
        var destination = RequirePath(context, DesImagePath, nameof(DesImagePath));

        using var image = ImageIO.Load(source);
        ImageIO.Save(image, destination, ImageIO.ToImageFormat(FileFormat));
    }
}
