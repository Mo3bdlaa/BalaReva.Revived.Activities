namespace BalaReva.EasyImage;

/// <summary>Output format for <see cref="ImageConverter"/>.</summary>
/// <remarks>
/// Member names and numeric values match the published package exactly; a workflow
/// persists the member name and Studio persists the value in some expression forms,
/// so neither may be renamed or renumbered.
/// </remarks>
public enum EnumFileFormat
{
    /// <summary>Windows bitmap.</summary>
    BMP = 1,
    /// <summary>Exchangeable Image File.</summary>
    EXIF = 2,
    /// <summary>Enhanced Metafile.</summary>
    EMF = 3,
    /// <summary>Graphics Interchange Format.</summary>
    GIF = 4,
    /// <summary>JPEG.</summary>
    JPEG = 5,
    /// <summary>Portable Network Graphics.</summary>
    PNG = 6,
    /// <summary>Tagged Image File Format.</summary>
    TIFF = 7,
    /// <summary>Windows Metafile.</summary>
    WMF = 8,
}
