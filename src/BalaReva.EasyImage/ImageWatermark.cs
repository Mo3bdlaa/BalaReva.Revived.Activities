using System.Activities;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Text;

namespace BalaReva.EasyImage;

/// <summary>Draws a text watermark onto an image.</summary>
[DisplayName("Image Watermark")]
[Description("Draws a text watermark onto an image and writes it to a destination folder.")]
public sealed class ImageWatermark : BaseWork
{
    /// <summary>Image to watermark.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Source File Path")]
    [Description("Full path of the image to watermark.")]
    public InArgument<string> SourceFilePath { get; set; } = null!;

    /// <summary>Folder the watermarked image is written into.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Destination Folder")]
    [Description("Folder to write the watermarked image into. Created if it does not exist.")]
    public InArgument<string> DestinationFolder { get; set; } = null!;

    /// <summary>Text to draw.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Text")]
    [Description("The watermark text.")]
    public InArgument<string> Text { get; set; } = null!;

    /// <summary>Font to draw the text in.</summary>
    [Category("Input")]
    [DisplayName("Font")]
    [Description("Font for the watermark text. Defaults to 24pt Arial.")]
    public InArgument<Font> Font { get; set; } = null!;

    /// <summary>Colour to draw the text in.</summary>
    [Category("Input")]
    [DisplayName("Fore Color")]
    [Description("Colour of the watermark text. Defaults to semi-transparent white.")]
    public InArgument<Color> ForeColor { get; set; } = null!;

    /// <summary>Where on the image to draw the text.</summary>
    [Category("Input")]
    [DisplayName("Text Position")]
    [Description("Top-left corner of the watermark text, in pixels.")]
    public InArgument<Point> TextPosition { get; set; } = null!;

    /// <summary>Format the watermarked image is written in.</summary>
    [Category("Input")]
    [DisplayName("Image Format")]
    [Description("Format to write the watermarked image in. Defaults to the source format.")]
    public InArgument<ImageFormat> ImageFormat { get; set; } = null!;

    /// <summary>Full path of the file that was written.</summary>
    [Category("Output")]
    [DisplayName("Output File")]
    [Description("Full path of the watermarked file that was written.")]
    public OutArgument<string> OutputFile { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context)
    {
        var source = RequireExistingFile(context, SourceFilePath, nameof(SourceFilePath));
        var folder = RequirePath(context, DestinationFolder, nameof(DestinationFolder));
        var text = Text?.Get(context);
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("Text is required.", nameof(Text));

        var format = ImageFormat?.Get(context) ?? ImageIO.FormatFromExtension(source);
        var position = TextPosition?.Get(context) ?? Point.Empty;
        var colour = ForeColor?.Get(context) ?? Color.FromArgb(128, Color.White);

        // Only dispose a font this activity created; one handed in by the workflow
        // belongs to the workflow and may be reused on the next iteration.
        var supplied = Font?.Get(context);
        var font = supplied ?? new Font("Arial", 24, FontStyle.Bold);

        try
        {
            using var image = ImageIO.Load(source);
            using (var graphics = Graphics.FromImage(image))
            {
                graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                using var brush = new SolidBrush(colour);
                graphics.DrawString(text, font, brush, position);
            }

            var output = Path.Combine(folder, Path.GetFileNameWithoutExtension(source) + Extension(format));
            ImageIO.Save(image, output, format);
            OutputFile.Set(context, output);
        }
        finally
        {
            if (supplied is null) font.Dispose();
        }
    }

    private static string Extension(ImageFormat format)
    {
        if (Equals(format, System.Drawing.Imaging.ImageFormat.Jpeg)) return ".jpg";
        if (Equals(format, System.Drawing.Imaging.ImageFormat.Bmp)) return ".bmp";
        if (Equals(format, System.Drawing.Imaging.ImageFormat.Gif)) return ".gif";
        if (Equals(format, System.Drawing.Imaging.ImageFormat.Tiff)) return ".tiff";
        if (Equals(format, System.Drawing.Imaging.ImageFormat.Emf)) return ".emf";
        if (Equals(format, System.Drawing.Imaging.ImageFormat.Wmf)) return ".wmf";
        return ".png";
    }
}
