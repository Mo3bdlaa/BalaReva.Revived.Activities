namespace BalaReva.Zip;

/// <summary>Reads and writes archives.</summary>
/// <remarks>
/// The activities talk to this rather than to the archive libraries directly, so argument
/// handling can be tested without one. Unlike the Office packages, the real implementation
/// here needs nothing installed, so it is tested too — against real archives, on every
/// agent.
/// </remarks>
public interface IArchiveService
{
    /// <summary>Extracts an archive into a folder.</summary>
    /// <returns>The full path of every file written.</returns>
    string[] Extract(ExtractRequest request);

    /// <summary>Writes files and folders into a zip archive.</summary>
    /// <returns>How many files were written.</returns>
    int Compress(CompressRequest request);
}

/// <summary>What <c>UnZipFile</c> asks the service to do.</summary>
/// <remarks>Not part of the published surface; it keeps the service signature readable.</remarks>
public sealed class ExtractRequest
{
    /// <summary>Full path of the archive to read.</summary>
    public string ArchiveFile { get; set; } = string.Empty;

    /// <summary>Folder to extract into. Created when it does not exist.</summary>
    public string DestinationFolder { get; set; } = string.Empty;

    /// <summary>Password, when the archive needs one.</summary>
    public string Password { get; set; } = string.Empty;

    /// <summary>Which extractor to use.</summary>
    public EnumExtractType ExtractType { get; set; } = EnumExtractType.Standard;

    /// <summary>
    /// Code page for entry names, for archives written by a tool that did not use UTF-8.
    /// Zero leaves the default alone.
    /// </summary>
    public int CodePage { get; set; }
}

/// <summary>What <c>ZipFilesCls</c> asks the service to do.</summary>
/// <remarks>Not part of the published surface; it keeps the service signature readable.</remarks>
public sealed class CompressRequest
{
    /// <summary>Full path of the archive to write.</summary>
    public string ArchiveFile { get; set; } = string.Empty;

    /// <summary>Files and folders to put in it. A folder goes in with its contents.</summary>
    public string[] Sources { get; set; } = [];

    /// <summary>Password to protect the archive with. Empty writes it unprotected.</summary>
    public string Password { get; set; } = string.Empty;
}
