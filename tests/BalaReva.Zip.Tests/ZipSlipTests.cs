using BalaReva.Zip;

namespace BalaReva.Zip.Tests;

/// <summary>
/// The reason this package was reimplemented. Every advisory against the published
/// package's dependencies was a path traversal on extraction, so these build archives
/// that actually contain the malicious entries and check that nothing is written outside
/// the extraction folder.
/// </summary>
public sealed class ZipSlipTests : IDisposable
{
    private readonly string _root =
        Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), Path.GetRandomFileName())).FullName;

    private readonly IArchiveService _service = new ArchiveService();

    private string Folder(string name) => Directory.CreateDirectory(Path.Combine(_root, name)).FullName;

    [Theory]
    [InlineData("../escaped.txt")]
    [InlineData("../../escaped.txt")]
    [InlineData("nested/../../escaped.txt")]
    [InlineData("..\\escaped.txt")]
    [InlineData("./../escaped.txt")]
    public void A_traversing_entry_is_refused_and_writes_nothing(string entryName)
    {
        var archive = WriteZip("evil.zip", (entryName, "owned"));
        var destination = Folder("out");

        Assert.Throws<UnauthorizedAccessException>(() => _service.Extract(new ExtractRequest
        {
            ArchiveFile = archive,
            DestinationFolder = destination,
        }));

        // The important half: not merely that it threw, but that nothing landed outside.
        Assert.False(File.Exists(Path.Combine(_root, "escaped.txt")));
        Assert.Empty(Directory.GetFiles(destination, "*", SearchOption.AllDirectories));
    }

    [Theory]
    [InlineData("/etc/passwd")]
    [InlineData("C:\\Windows\\System32\\evil.dll")]
    [InlineData("\\\\server\\share\\evil.dll")]
    public void An_absolute_entry_is_refused(string entryName)
    {
        var archive = WriteZip("absolute.zip", (entryName, "owned"));

        Assert.Throws<UnauthorizedAccessException>(() => _service.Extract(new ExtractRequest
        {
            ArchiveFile = archive,
            DestinationFolder = Folder("out"),
        }));
    }

    [Fact]
    public void A_sibling_folder_sharing_a_name_prefix_is_still_outside()
    {
        // The check compares against the folder plus a separator for exactly this: without
        // it, extracting into "out" would accept an entry resolving into "outside".
        Directory.CreateDirectory(Path.Combine(_root, "outside"));
        var archive = WriteZip("prefix.zip", ("../outside/evil.txt", "owned"));

        Assert.Throws<UnauthorizedAccessException>(() => _service.Extract(new ExtractRequest
        {
            ArchiveFile = archive,
            DestinationFolder = Folder("out"),
        }));

        Assert.False(File.Exists(Path.Combine(_root, "outside", "evil.txt")));
    }

    [Fact]
    public void An_honest_archive_still_extracts()
    {
        var archive = WriteZip("good.zip",
            ("readme.txt", "hello"),
            ("nested/deep/file.txt", "world"));

        var destination = Folder("out");
        var written = _service.Extract(new ExtractRequest
        {
            ArchiveFile = archive,
            DestinationFolder = destination,
        });

        Assert.Equal(2, written.Length);
        Assert.Equal("hello", File.ReadAllText(Path.Combine(destination, "readme.txt")));
        Assert.Equal("world", File.ReadAllText(Path.Combine(destination, "nested", "deep", "file.txt")));
    }

    [Fact]
    public void A_dot_dot_inside_the_path_that_stays_within_bounds_is_allowed()
    {
        // "a/../b.txt" resolves to "b.txt", which is inside the folder. Refusing every
        // entry containing ".." would be the lazy check and would break honest archives.
        var archive = WriteZip("winding.zip", ("a/../b.txt", "fine"));
        var destination = Folder("out");

        _service.Extract(new ExtractRequest { ArchiveFile = archive, DestinationFolder = destination });

        Assert.Equal("fine", File.ReadAllText(Path.Combine(destination, "b.txt")));
    }

    /// <summary>Writes a zip whose entry names reach the extractor exactly as given.</summary>
    private string WriteZip(string name, params (string Entry, string Content)[] entries) =>
        HostileZip.Write(Path.Combine(_root, name), entries);

    public void Dispose()
    {
        try { Directory.Delete(_root, recursive: true); }
        catch (IOException) { /* the agent will clean its own temp */ }
    }
}
