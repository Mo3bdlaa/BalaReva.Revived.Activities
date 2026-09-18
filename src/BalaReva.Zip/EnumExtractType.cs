namespace BalaReva.Zip;

/// <summary>
/// Which extractor <c>UnZipFile</c> should use.
/// </summary>
/// <remarks>
/// Mirrors the published <c>BalaReva.Zip.EnumExtractType</c>. A workflow persists the
/// member name, so neither name may change.
/// </remarks>
public enum EnumExtractType
{
    /// <summary>Zip archives, including password-protected ones.</summary>
    Standard = 0,

    /// <summary>
    /// Anything else the archive readers recognise: rar, 7z, tar, gzip, bzip2 and the
    /// rest.
    /// </summary>
    UniExtract = 1,
}
