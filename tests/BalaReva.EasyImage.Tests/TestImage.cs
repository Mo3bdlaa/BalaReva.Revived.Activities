using System.Activities;
using System.Drawing;
using System.Drawing.Imaging;

namespace BalaReva.EasyImage.Tests;

/// <summary>A throwaway folder with images in it, cleaned up on dispose.</summary>
public sealed class TestImages : IDisposable
{
    public TestImages()
    {
        Folder = Path.Combine(Path.GetTempPath(), $"easyimage-{Guid.NewGuid():N}");
        Directory.CreateDirectory(Folder);
    }

    public string Folder { get; }

    /// <summary>Writes a solid-colour image and returns its path.</summary>
    public string Create(string name, int width, int height, Color? fill = null)
    {
        var path = At(name);
        using var bitmap = new Bitmap(width, height);
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(fill ?? Color.CornflowerBlue);
        }
        bitmap.Save(path, FormatFor(name));
        return path;
    }

    /// <summary>A path inside the folder that nothing has been written to yet.</summary>
    public string At(string name) => Path.Combine(Folder, name);

    /// <summary>Pixel dimensions of an image on disk.</summary>
    public static (int Width, int Height) SizeOf(string path)
    {
        using var stream = File.OpenRead(path);
        using var image = Image.FromStream(stream, false, true);
        return (image.Width, image.Height);
    }

    /// <summary>The format GDI+ reports for an image on disk.</summary>
    public static ImageFormat FormatOf(string path)
    {
        using var stream = File.OpenRead(path);
        using var image = Image.FromStream(stream, false, true);
        return image.RawFormat;
    }

    /// <summary>Fills in the arguments every image activity requires.</summary>
    public static T Ready<T>(T activity) where T : BaseWork
    {
        activity.ContinueOnError ??= new InArgument<bool>(false);
        activity.Delay ??= new InArgument<short>(0);
        activity.ExecutionResult ??= new OutArgument<bool>();
        return activity;
    }

    /// <summary>Invokes an activity as a workflow root and returns its outputs.</summary>
    public static IDictionary<string, object> Run(BaseWork activity) =>
        WorkflowInvoker.Invoke(Ready(activity));

    private static ImageFormat FormatFor(string name) =>
        Path.GetExtension(name).ToLowerInvariant() switch
        {
            ".jpg" or ".jpeg" => ImageFormat.Jpeg,
            ".bmp" => ImageFormat.Bmp,
            ".gif" => ImageFormat.Gif,
            _ => ImageFormat.Png,
        };

    public void Dispose()
    {
        if (Directory.Exists(Folder)) Directory.Delete(Folder, recursive: true);
    }
}
