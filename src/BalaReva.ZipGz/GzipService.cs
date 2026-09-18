using BalaReva.Revived.Shared;
using ICSharpCode.SharpZipLib.GZip;
using ICSharpCode.SharpZipLib.Tar;

namespace BalaReva.ZipGz;

/// <summary>Extracts gzip and tar archives with SharpZipLib.</summary>
/// <remarks>
/// Every write goes through <see cref="SafeExtractPath"/> first. SharpZipLib 1.2.0, which
/// the published package both depended on and vendored, was vulnerable to exactly the
/// traversal that check refuses: CVE-2021-32840 covered tar, CVE-2021-32842 zip.
/// </remarks>
public sealed class GzipService : IGzipService
{
    /// <summary>The service the activity uses when no extension is registered.</summary>
    public static IGzipService Instance { get; } = new GzipService();

    /// <inheritdoc />
    public string[] Extract(string archiveFile, string destinationFolder)
    {
        if (!File.Exists(archiveFile))
            throw new FileNotFoundException($"'{archiveFile}' does not exist.", archiveFile);

        Directory.CreateDirectory(destinationFolder);

        return IsGzip(archiveFile)
            ? ExtractGzip(archiveFile, destinationFolder)
            : ExtractTar(File.OpenRead(archiveFile), destinationFolder);
    }

    /// <summary>
    /// A .tar.gz is a tar inside a gzip, and a plain .gz is one file inside a gzip. Both
    /// arrive here, so the decompressed stream decides which it was.
    /// </summary>
    private static string[] ExtractGzip(string archiveFile, string destinationFolder)
    {
        using var decompressed = new MemoryStream();
        using (var source = File.OpenRead(archiveFile))
        using (var gzip = new GZipInputStream(source))
        {
            gzip.CopyTo(decompressed);
        }
        decompressed.Position = 0;

        if (LooksLikeTar(decompressed))
        {
            decompressed.Position = 0;
            return ExtractTar(decompressed, destinationFolder, leaveOpen: true);
        }

        // A single compressed file: gzip carries no entry name of its own that can be
        // trusted, so it takes the archive's name without its .gz suffix.
        var name = Path.GetFileNameWithoutExtension(archiveFile);
        if (string.IsNullOrWhiteSpace(name)) name = "extracted";

        var target = SafeExtractPath.Resolve(destinationFolder, name);
        SafeExtractPath.EnsureFolder(target);

        decompressed.Position = 0;
        using (var destination = File.Create(target))
        {
            decompressed.CopyTo(destination);
        }
        return [target];
    }

    private static string[] ExtractTar(Stream source, string destinationFolder, bool leaveOpen = false)
    {
        var written = new List<string>();

        // TarInputStream rather than TarArchive.ExtractContents: the latter is what the
        // advisories were written against, because it joins the entry name to the
        // destination itself. Reading entry by entry is what makes the path check
        // possible.
        using var tar = new TarInputStream(source, System.Text.Encoding.UTF8) { IsStreamOwner = !leaveOpen };

        while (tar.GetNextEntry() is { } entry)
        {
            var target = SafeExtractPath.Resolve(destinationFolder, entry.Name);

            if (entry.IsDirectory)
            {
                Directory.CreateDirectory(target);
                continue;
            }

            SafeExtractPath.EnsureFolder(target);
            using (var destination = File.Create(target))
            {
                tar.CopyEntryContents(destination);
            }
            written.Add(target);
        }
        return [.. written];
    }

    /// <summary>Whether a file starts with the gzip magic number.</summary>
    private static bool IsGzip(string path)
    {
        using var stream = File.OpenRead(path);
        return stream.ReadByte() == 0x1F && stream.ReadByte() == 0x8B;
    }

    /// <summary>
    /// Whether a decompressed stream is a tar: the format writes "ustar" at offset 257 of
    /// its first header block.
    /// </summary>
    private static bool LooksLikeTar(Stream stream)
    {
        if (stream.Length < 265) return false;

        stream.Position = 257;
        var magic = new byte[5];
        return stream.ReadAtLeast(magic, magic.Length, throwOnEndOfStream: false) == magic.Length
            && System.Text.Encoding.ASCII.GetString(magic) == "ustar";
    }
}
