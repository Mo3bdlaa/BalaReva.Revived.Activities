namespace BalaReva.Excel.Utilities;

/// <summary>Position and size of an object placed on a sheet, in points.</summary>
public interface IObjectSize
{
    /// <summary>Distance from the left edge of the sheet.</summary>
    double Left { get; set; }

    /// <summary>Distance from the top edge of the sheet.</summary>
    double Top { get; set; }

    /// <summary>Width.</summary>
    double Width { get; set; }

    /// <summary>Height.</summary>
    double Height { get; set; }
}

/// <summary>Where and how big to place an image on a sheet, in points.</summary>
public sealed class ObjectSize : IObjectSize
{
    /// <inheritdoc />
    public double Left { get; set; }

    /// <inheritdoc />
    public double Top { get; set; }

    /// <inheritdoc />
    public double Width { get; set; }

    /// <inheritdoc />
    public double Height { get; set; }
}
