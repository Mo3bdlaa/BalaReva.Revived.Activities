using System.Drawing;
using System.Drawing.Drawing2D;

namespace BalaReva.EasyImage;

/// <summary>Lays two images out into one canvas.</summary>
internal static class ImageLayout
{
    /// <summary>
    /// Writes <paramref name="first"/> and <paramref name="second"/> into one image,
    /// side by side when <paramref name="horizontal"/> is true and stacked otherwise.
    /// </summary>
    /// <remarks>
    /// The canvas takes the larger of the two sizes on the cross axis, so images of
    /// different sizes are padded rather than stretched or clipped.
    /// </remarks>
    public static void Join(string first, string second, string destination, bool horizontal)
    {
        using var left = ImageIO.Load(first);
        using var right = ImageIO.Load(second);

        var width = horizontal ? left.Width + right.Width : Math.Max(left.Width, right.Width);
        var height = horizontal ? Math.Max(left.Height, right.Height) : left.Height + right.Height;

        using var canvas = new Bitmap(width, height);
        using (var graphics = Graphics.FromImage(canvas))
        {
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.Clear(Color.Transparent);
            graphics.DrawImage(left, 0, 0, left.Width, left.Height);
            if (horizontal) graphics.DrawImage(right, left.Width, 0, right.Width, right.Height);
            else graphics.DrawImage(right, 0, left.Height, right.Width, right.Height);
        }

        ImageIO.Save(canvas, destination, ImageIO.FormatFromExtension(destination));
    }
}
