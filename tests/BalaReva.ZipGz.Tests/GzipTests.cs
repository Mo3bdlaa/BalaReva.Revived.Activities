using System.Activities;
using System.Text;
using ICSharpCode.SharpZipLib.GZip;
using ICSharpCode.SharpZipLib.Tar;

namespace BalaReva.ZipGz.Tests;

/// <summary>A stand-in service that records what it was asked to do.</summary>
public sealed class FakeGzipService : IGzipService
{
    public string? LastArchive { get; private set; }

    public string? LastDestination { get; private set; }

    public string[] Extract(string archiveFile, string destinationFolder)
    {
        LastArchive = archiveFile;
        LastDestination = destinationFolder;
        return [];
    }
}

/// <summary>
/// The tar path traversal CVE-2021-32840 covered, and the formats the activity accepts.
/// </summary>
public sealed class GzipTests : IDisposable
{
    private readonly string _root =
        Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), Path.GetRandomFileName())).FullName;

    private readonly IGzipService _service = new GzipService();

    private string Folder(string name) => Directory.CreateDirectory(Path.Combine(_root, name)).FullName;

    [Theory]
    [InlineData("../escaped.txt")]
    [InlineData("../../escaped.txt")]
    [InlineData("nested/../../escaped.txt")]
    [InlineData("/etc/passwd")]
    public void A_traversing_tar_entry_is_refused_and_writes_nothing(string entryName)
    {
        var archive = WriteTar("evil.tar", (entryName, "owned"));
        var destination = Folder("out");

        Assert.Throws<UnauthorizedAccessException>(() => _service.Extract(archive, destination));

        Assert.False(File.Exists(Path.Combine(_root, "escaped.txt")));
        Assert.Empty(Directory.GetFiles(destination, "*", SearchOption.AllDirectories));
    }

    [Fact]
    public void A_traversing_entry_inside_a_gzipped_tar_is_refused_too()
    {
        // The .tar.gz path decompresses first, so the check has to sit on the tar reader
        // rather than on the file being opened.
        var archive = Gzip(WriteTar("evil.tar", ("../escaped.txt", "owned")), "evil.tar.gz");

        Assert.Throws<UnauthorizedAccessException>(() => _service.Extract(archive, Folder("out")));
        Assert.False(File.Exists(Path.Combine(_root, "escaped.txt")));
    }

    [Fact]
    public void An_honest_tar_extracts()
    {
        var archive = WriteTar("good.tar", ("readme.txt", "hello"), ("nested/deep.txt", "world"));
        var destination = Folder("out");

        var written = _service.Extract(archive, destination);

        Assert.Equal(2, written.Length);
        Assert.Equal("hello", File.ReadAllText(Path.Combine(destination, "readme.txt")));
        Assert.Equal("world", File.ReadAllText(Path.Combine(destination, "nested", "deep.txt")));
    }

    [Fact]
    public void A_gzipped_tar_extracts()
    {
        var archive = Gzip(WriteTar("good.tar", ("readme.txt", "hello")), "good.tar.gz");
        var destination = Folder("out");

        _service.Extract(archive, destination);

        Assert.Equal("hello", File.ReadAllText(Path.Combine(destination, "readme.txt")));
    }

    [Fact]
    public void A_gzip_holding_a_single_file_extracts_under_the_archive_name()
    {
        // A plain .gz carries no entry name worth trusting, so the file takes the
        // archive's name without its suffix.
        var plain = Path.Combine(_root, "notes.txt");
        File.WriteAllText(plain, "just one file");
        var archive = Gzip(plain, "notes.txt.gz");

        var destination = Folder("out");
        var written = _service.Extract(archive, destination);

        var target = Assert.Single(written);
        Assert.Equal(Path.Combine(destination, "notes.txt"), target);
        Assert.Equal("just one file", File.ReadAllText(target));
    }

    [Fact]
    public void The_format_is_decided_by_the_bytes_not_the_extension()
    {
        // A tar named .gz is still a tar, and the reader should say so rather than
        // failing on the gzip magic number.
        var mislabelled = Path.Combine(_root, "actually.tar.gz");
        File.Copy(WriteTar("real.tar", ("readme.txt", "hello")), mislabelled);

        var destination = Folder("out");
        _service.Extract(mislabelled, destination);

        Assert.Equal("hello", File.ReadAllText(Path.Combine(destination, "readme.txt")));
    }

    [Fact]
    public void An_archive_that_is_not_there_says_so() =>
        Assert.Throws<FileNotFoundException>(
            () => _service.Extract(Path.Combine(_root, "missing.tar"), _root));

    [Fact]
    public void The_activity_forwards_both_of_its_arguments()
    {
        var service = new FakeGzipService();
        var invoker = new WorkflowInvoker(new UnZipCls
        {
            ZipFile = new InArgument<string>(@"C:\in\archive.tar.gz"),
            ExtractFolderPath = new InArgument<string>(@"C:\out"),
        });
        invoker.Extensions.Add(service);
        invoker.Invoke();

        Assert.Equal(@"C:\in\archive.tar.gz", service.LastArchive);
        Assert.Equal(@"C:\out", service.LastDestination);
    }

    [Theory]
    [InlineData(null, "out")]
    [InlineData("", "out")]
    [InlineData("a.tar", null)]
    [InlineData("a.tar", "")]
    public void The_activity_rejects_a_missing_path(string? archive, string? destination)
    {
        var invoker = new WorkflowInvoker(new UnZipCls
        {
            ZipFile = new InArgument<string>(archive!),
            ExtractFolderPath = new InArgument<string>(destination!),
        });
        invoker.Extensions.Add(new FakeGzipService());

        Assert.ThrowsAny<Exception>(invoker.Invoke);
    }

    /// <summary>
    /// Writes a tar holding the given entries, with the names untouched.
    /// </summary>
    /// <remarks>
    /// TarOutputStream does not sanitise the way the zip writer does, so unlike the zip
    /// tests this needs no hand-assembled archive.
    /// </remarks>
    private string WriteTar(string name, params (string Entry, string Content)[] entries)
    {
        var path = Path.Combine(_root, name);
        using var stream = new TarOutputStream(File.Create(path), Encoding.UTF8);

        foreach (var (entry, content) in entries)
        {
            var bytes = Encoding.UTF8.GetBytes(content);
            var header = TarEntry.CreateTarEntry(entry);
            header.Size = bytes.Length;

            stream.PutNextEntry(header);
            stream.Write(bytes, 0, bytes.Length);
            stream.CloseEntry();
        }
        return path;
    }

    private string Gzip(string source, string name)
    {
        var path = Path.Combine(_root, name);
        using (var input = File.OpenRead(source))
        using (var output = new GZipOutputStream(File.Create(path)))
        {
            input.CopyTo(output);
        }
        return path;
    }

    public void Dispose()
    {
        try { Directory.Delete(_root, recursive: true); }
        catch (IOException) { /* the agent will clean its own temp */ }
    }
}
