using System.Drawing;
using System.Drawing.Imaging;

namespace BalaReva.EasyImage;

/// <summary>Loading, saving and format mapping shared by the image activities.</summary>
internal static class ImageIO
{
    /// <summary>
    /// Loads a bitmap without keeping the source file locked.
    /// </summary>
    /// <remarks>
    /// <see cref="Image.FromFile(string)"/> keeps a lock on the file for the lifetime
    /// of the image, which makes writing the result back to the same path fail. Copying
    /// through a stream avoids that, and it is the difference between an activity that
    /// can overwrite its input and one that cannot.
    /// </remarks>
    public static Bitmap Load(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var loaded = Image.FromStream(stream, useEmbeddedColorManagement: false, validateImageData: true);
        return new Bitmap(loaded);
    }

    /// <summary>Maps the package's format enum onto GDI+.</summary>
    public static ImageFormat ToImageFormat(EnumFileFormat format) => format switch
    {
        EnumFileFormat.BMP => ImageFormat.Bmp,
        EnumFileFormat.EXIF => ImageFormat.Exif,
        EnumFileFormat.EMF => ImageFormat.Emf,
        EnumFileFormat.GIF => ImageFormat.Gif,
        EnumFileFormat.JPEG => ImageFormat.Jpeg,
        EnumFileFormat.PNG => ImageFormat.Png,
        EnumFileFormat.TIFF => ImageFormat.Tiff,
        EnumFileFormat.WMF => ImageFormat.Wmf,
        _ => throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported image format."),
    };

    /// <summary>Saves <paramref name="image"/>, creating the destination folder if needed.</summary>
    public static void Save(Image image, string path, ImageFormat format)
    {
        EnsureFolder(path);
        // Buffer first: the destination may be the source, which is still open.
        using var buffer = new MemoryStream();
        image.Save(buffer, format);
        File.WriteAllBytes(path, buffer.ToArray());
    }

    /// <summary>Saves as JPEG at the given quality, 0 to 100.</summary>
    public static void SaveJpeg(Image image, string path, int quality)
    {
        if (quality is < 0 or > 100)
        {
            throw new ArgumentOutOfRangeException(
                nameof(quality), quality, "Image quality must be between 0 and 100.");
        }

        var codec = ImageCodecInfo.GetImageEncoders()
            .FirstOrDefault(c => c.FormatID == ImageFormat.Jpeg.Guid)
            ?? throw new InvalidOperationException("No JPEG encoder is available on this machine.");

        using var parameters = new EncoderParameters(1);
        using var parameter = new EncoderParameter(Encoder.Quality, (long)quality);
        parameters.Param[0] = parameter;

        EnsureFolder(path);
        using var buffer = new MemoryStream();
        image.Save(buffer, codec, parameters);
        File.WriteAllBytes(path, buffer.ToArray());
    }

    /// <summary>Picks an output format from the destination file's extension.</summary>
    public static ImageFormat FormatFromExtension(string path) =>
        Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".bmp" => ImageFormat.Bmp,
            ".gif" => ImageFormat.Gif,
            ".jpg" or ".jpeg" => ImageFormat.Jpeg,
            ".tif" or ".tiff" => ImageFormat.Tiff,
            ".emf" => ImageFormat.Emf,
            ".wmf" => ImageFormat.Wmf,
            _ => ImageFormat.Png,
        };

    private static void EnsureFolder(string path)
    {
        var folder = Path.GetDirectoryName(Path.GetFullPath(path));
        if (!string.IsNullOrEmpty(folder)) Directory.CreateDirectory(folder);
    }
}
