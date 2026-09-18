using System.Activities;
using BalaReva.Zip;
using BalaReva.ZipUnzip;

namespace BalaReva.Zip.Tests;

/// <summary>A stand-in archive service that records what it was asked to do.</summary>
public sealed class FakeArchiveService : IArchiveService
{
    public ExtractRequest? LastExtract { get; private set; }

    public CompressRequest? LastCompress { get; private set; }

    public string[] Extract(ExtractRequest request)
    {
        LastExtract = request;
        return [];
    }

    public int Compress(CompressRequest request)
    {
        LastCompress = request;
        return request.Sources.Length;
    }
}

/// <summary>
/// Argument handling for the two activities, and a real round trip through the service
/// underneath them.
/// </summary>
public sealed class ActivityTests : IDisposable
{
    private readonly string _root =
        Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), Path.GetRandomFileName())).FullName;

    [Fact]
    public void UnZipFile_gathers_its_arguments_into_one_request()
    {
        var service = new FakeArchiveService();
        Run(
            new UnZipFile
            {
                strZipFile = new InArgument<string>(@"C:\in\archive.rar"),
                ExtractFolderPath = new InArgument<string>(@"C:\out"),
                Password = new InArgument<string>("secret"),
                ExtractType = EnumExtractType.UniExtract,
                CodePage = new InArgument<int>(866),
            },
            service);

        var request = Assert.IsType<ExtractRequest>(service.LastExtract);
        Assert.Equal(@"C:\in\archive.rar", request.ArchiveFile);
        Assert.Equal(@"C:\out", request.DestinationFolder);
        Assert.Equal("secret", request.Password);
        Assert.Equal(EnumExtractType.UniExtract, request.ExtractType);
        Assert.Equal(866, request.CodePage);
    }

    [Fact]
    public void UnZipFile_defaults_to_the_standard_zip_reader()
    {
        var service = new FakeArchiveService();
        Run(
            new UnZipFile
            {
                strZipFile = new InArgument<string>("a.zip"),
                ExtractFolderPath = new InArgument<string>("out"),
            },
            service);

        Assert.Equal(EnumExtractType.Standard, service.LastExtract!.ExtractType);
        Assert.Equal(string.Empty, service.LastExtract.Password);
        Assert.Equal(0, service.LastExtract.CodePage);
    }

    [Theory]
    [InlineData(null, "out")]
    [InlineData("", "out")]
    [InlineData("a.zip", null)]
    [InlineData("a.zip", "")]
    public void UnZipFile_rejects_a_missing_path(string? archive, string? destination)
    {
        var activity = new UnZipFile
        {
            strZipFile = new InArgument<string>(archive!),
            ExtractFolderPath = new InArgument<string>(destination!),
        };

        Assert.ThrowsAny<Exception>(() => Run(activity, new FakeArchiveService()));
    }

    [Fact]
    public void ZipFilesCls_gathers_its_arguments_into_one_request()
    {
        var service = new FakeArchiveService();
        Run(
            new ZipFilesCls
            {
                ZipFile = new InArgument<string>(@"C:\out\archive.zip"),
                FolderFilesPath = new InArgument<string[]>(_ => new[] { @"C:\a.txt", @"C:\folder" }),
                Password = new InArgument<string>("secret"),
            },
            service);

        var request = Assert.IsType<CompressRequest>(service.LastCompress);
        Assert.Equal(@"C:\out\archive.zip", request.ArchiveFile);
        Assert.Equal(new[] { @"C:\a.txt", @"C:\folder" }, request.Sources);
        Assert.Equal("secret", request.Password);
    }

    [Fact]
    public void ZipFilesCls_rejects_an_empty_source_list()
    {
        var activity = new ZipFilesCls
        {
            ZipFile = new InArgument<string>("a.zip"),
            FolderFilesPath = new InArgument<string[]>(_ => Array.Empty<string>()),
        };

        Assert.ThrowsAny<Exception>(() => Run(activity, new FakeArchiveService()));
    }

    [Fact]
    public void A_file_and_a_folder_survive_a_round_trip()
    {
        var source = Directory.CreateDirectory(Path.Combine(_root, "src")).FullName;
        File.WriteAllText(Path.Combine(source, "top.txt"), "top");
        var nested = Directory.CreateDirectory(Path.Combine(source, "nested")).FullName;
        File.WriteAllText(Path.Combine(nested, "deep.txt"), "deep");

        var loose = Path.Combine(_root, "loose.txt");
        File.WriteAllText(loose, "loose");

        var archive = Path.Combine(_root, "out.zip");
        var service = new ArchiveService();

        var written = service.Compress(new CompressRequest
        {
            ArchiveFile = archive,
            Sources = [source, loose],
        });
        Assert.Equal(3, written);

        var destination = Directory.CreateDirectory(Path.Combine(_root, "back")).FullName;
        service.Extract(new ExtractRequest { ArchiveFile = archive, DestinationFolder = destination });

        // A folder goes in under its own name, so unzipping reproduces it rather than
        // spilling its contents into the destination.
        Assert.Equal("top", File.ReadAllText(Path.Combine(destination, "src", "top.txt")));
        Assert.Equal("deep", File.ReadAllText(Path.Combine(destination, "src", "nested", "deep.txt")));
        Assert.Equal("loose", File.ReadAllText(Path.Combine(destination, "loose.txt")));
    }

    [Fact]
    public void A_password_protected_archive_survives_a_round_trip()
    {
        var file = Path.Combine(_root, "secret.txt");
        File.WriteAllText(file, "classified");

        var archive = Path.Combine(_root, "locked.zip");
        var service = new ArchiveService();

        service.Compress(new CompressRequest
        {
            ArchiveFile = archive,
            Sources = [file],
            Password = "hunter2",
        });

        var destination = Directory.CreateDirectory(Path.Combine(_root, "back")).FullName;
        service.Extract(new ExtractRequest
        {
            ArchiveFile = archive,
            DestinationFolder = destination,
            Password = "hunter2",
        });

        Assert.Equal("classified", File.ReadAllText(Path.Combine(destination, "secret.txt")));
    }

    [Fact]
    public void The_wrong_password_does_not_quietly_produce_rubbish()
    {
        var file = Path.Combine(_root, "secret.txt");
        File.WriteAllText(file, "classified");

        var archive = Path.Combine(_root, "locked.zip");
        var service = new ArchiveService();
        service.Compress(new CompressRequest
        {
            ArchiveFile = archive,
            Sources = [file],
            Password = "hunter2",
        });

        Assert.ThrowsAny<Exception>(() => service.Extract(new ExtractRequest
        {
            ArchiveFile = archive,
            DestinationFolder = Directory.CreateDirectory(Path.Combine(_root, "back")).FullName,
            Password = "wrong",
        }));
    }

    [Fact]
    public void An_archive_that_is_not_there_says_so() =>
        Assert.Throws<FileNotFoundException>(() => new ArchiveService().Extract(new ExtractRequest
        {
            ArchiveFile = Path.Combine(_root, "missing.zip"),
            DestinationFolder = _root,
        }));

    [Fact]
    public void Compressing_something_that_is_not_there_says_so() =>
        Assert.Throws<FileNotFoundException>(() => new ArchiveService().Compress(new CompressRequest
        {
            ArchiveFile = Path.Combine(_root, "out.zip"),
            Sources = [Path.Combine(_root, "missing.txt")],
        }));

    private static void Run(Activity activity, IArchiveService service)
    {
        var invoker = new WorkflowInvoker(activity);
        invoker.Extensions.Add(service);
        invoker.Invoke();
    }

    public void Dispose()
    {
        try { Directory.Delete(_root, recursive: true); }
        catch (IOException) { /* the agent will clean its own temp */ }
    }
}
