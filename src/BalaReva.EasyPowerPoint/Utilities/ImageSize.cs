namespace BalaReva.EasyPowerPoint.Utilities;

/// <summary>Where a picture sits on a slide, and how big it is, all in points.</summary>
public sealed class ImageSize
{
    /// <summary>Height in points.</summary>
    public float Height { get; set; }

    /// <summary>Distance from the left edge of the slide, in points.</summary>
    public float Left { get; set; }

    /// <summary>Distance from the top edge of the slide, in points.</summary>
    public float Top { get; set; }

    /// <summary>Width in points.</summary>
    public float Width { get; set; }
}
