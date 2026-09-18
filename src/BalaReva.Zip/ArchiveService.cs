using System.Text;
using BalaReva.Revived.Shared;
using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;
using SharpCompress.Archives;
using SharpCompress.Common;
using SharpCompress.Readers;

namespace BalaReva.Zip;

/// <summary>Reads and writes archives with SharpZipLib and SharpCompress.</summary>
/// <remarks>
/// Every write goes through <see cref="SafeExtractPath"/> first. See the note there for
/// why that check is made here rather than trusted to the libraries.
/// </remarks>
public sealed class ArchiveService : IArchiveService
{
    /// <summary>The service the activities use when no extension is registered.</summary>
    public static IArchiveService Instance { get; } = new ArchiveService();

    static ArchiveService() =>
        // Needed for the legacy code pages a CodePage argument would ask for; .NET Core
        // ships only UTF-8 and Latin-1 without it.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    /// <inheritdoc />
    public string[] Extract(ExtractRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (!File.Exists(request.ArchiveFile))
            throw new FileNotFoundException($"'{request.ArchiveFile}' does not exist.", request.ArchiveFile);

        Directory.CreateDirectory(request.DestinationFolder);

        return request.ExtractType == EnumExtractType.UniExtract
            ? ExtractAnyFormat(request)
            : ExtractZip(request);
    }

    /// <summary>Extracts a zip, which is the only format that takes a password here.</summary>
    private static string[] ExtractZip(ExtractRequest request)
    {
        var written = new List<string>();

        using var file = new ZipFile(File.OpenRead(request.ArchiveFile))
        {
            IsStreamOwner = true,
        };

        // Per-archive rather than the process-wide ZipStrings.CodePage the old version
        // used, which was a global that leaked between concurrent extractions.
        if (request.CodePage > 0)
            file.StringCodec = StringCodec.FromCodePage(request.CodePage);

        if (request.Password.Length > 0) file.Password = request.Password;

        foreach (ZipEntry entry in file)
        {
            var target = SafeExtractPath.Resolve(request.DestinationFolder, entry.Name);

            if (entry.IsDirectory)
            {
                Directory.CreateDirectory(target);
                continue;
            }
            if (!entry.IsFile) continue;

            SafeExtractPath.EnsureFolder(target);
            using (var source = file.GetInputStream(entry))
            using (var destination = File.Create(target))
            {
                source.CopyTo(destination);
            }
            written.Add(target);
        }
        return [.. written];
    }

    /// <summary>Extracts rar, 7z, tar, gzip, bzip2 and the other formats.</summary>
    private static string[] ExtractAnyFormat(ExtractRequest request)
    {
        var written = new List<string>();
        var options = new ReaderOptions
        {
            Password = request.Password.Length > 0 ? request.Password : null,
        };

        using var archive = ArchiveFactory.Open(request.ArchiveFile, options);
        foreach (var entry in archive.Entries)
        {
            if (entry.Key is null) continue;

            var target = SafeExtractPath.Resolve(request.DestinationFolder, entry.Key);

            if (entry.IsDirectory)
            {
                Directory.CreateDirectory(target);
                continue;
            }

            SafeExtractPath.EnsureFolder(target);
            using (var source = entry.OpenEntryStream())
            using (var destination = File.Create(target))
            {
                source.CopyTo(destination);
            }
            written.Add(target);
        }
        return [.. written];
    }

    /// <inheritdoc />
    public int Compress(CompressRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (string.IsNullOrWhiteSpace(request.ArchiveFile))
            throw new ArgumentException("ZipFile is required.", nameof(request));
        if (request.Sources.Length == 0)
            throw new ArgumentException("FolderFilesPath is required.", nameof(request));

        var folder = Path.GetDirectoryName(Path.GetFullPath(request.ArchiveFile));
        if (!string.IsNullOrEmpty(folder)) Directory.CreateDirectory(folder);

        using var stream = new ZipOutputStream(File.Create(request.ArchiveFile));
        if (request.Password.Length > 0) stream.Password = request.Password;

        var count = 0;
        foreach (var source in request.Sources)
        {
            if (string.IsNullOrWhiteSpace(source)) continue;

            if (Directory.Exists(source))
            {
                // A folder goes in under its own name, so unzipping reproduces it rather
                // than spilling its contents into the destination.
                var root = Path.TrimEndingDirectorySeparator(Path.GetFullPath(source));
                var prefix = Path.GetFileName(root);

                foreach (var file in Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories))
                {
                    var relative = Path.GetRelativePath(root, file).Replace('\\', '/');
                    Write(stream, file, $"{prefix}/{relative}");
                    count++;
                }
            }
            else if (File.Exists(source))
            {
                Write(stream, source, Path.GetFileName(source));
                count++;
            }
            else
            {
                throw new FileNotFoundException($"'{source}' does not exist.", source);
            }
        }

        stream.Finish();
        return count;
    }

    private static void Write(ZipOutputStream stream, string file, string entryName)
    {
        var info = new FileInfo(file);
        stream.PutNextEntry(new ZipEntry(entryName)
        {
            DateTime = info.LastWriteTime,
            Size = info.Length,
        });

        using var source = File.OpenRead(file);
        source.CopyTo(stream);
        stream.CloseEntry();
    }
}
