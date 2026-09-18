namespace BalaReva.ZipGz;

/// <summary>Extracts gzip and tar archives.</summary>
/// <remarks>
/// The activity talks to this rather than to SharpZipLib directly, so argument handling
/// can be tested without one. The real implementation needs nothing installed, so it is
/// tested too, against real archives.
/// </remarks>
public interface IGzipService
{
    /// <summary>Extracts an archive into a folder.</summary>
    /// <returns>The full path of every file written.</returns>
    string[] Extract(string archiveFile, string destinationFolder);
}
