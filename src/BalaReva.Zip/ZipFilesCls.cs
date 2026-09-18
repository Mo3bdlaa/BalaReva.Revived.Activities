using System.Activities;
using System.ComponentModel;

namespace BalaReva.ZipUnzip;

/// <summary>Writes files and folders into a zip archive.</summary>
/// <remarks>
/// Note the <c>Cls</c> suffix and the <c>ZipFile</c> argument, both as the published
/// package had them.
/// </remarks>
[DisplayName("Zip Files")]
[Description("Writes files and folders into a zip archive.")]
public sealed class ZipFilesCls : CodeActivity
{
    /// <summary>Files and folders to put in the archive.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Folder Files Path")]
    [Description("Files and folders to put in the archive. A folder goes in with its contents.")]
    public InArgument<string[]> FolderFilesPath { get; set; } = null!;

    /// <summary>Full path of the archive to write.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Zip File")]
    [Description("Full path of the archive to write.")]
    public InArgument<string> ZipFile { get; set; } = null!;

    /// <summary>Password to protect the archive with. Empty writes it unprotected.</summary>
    [Category("Input")]
    [DisplayName("Password")]
    [Description("Password to protect the archive with. Empty writes it unprotected.")]
    public InArgument<string> Password { get; set; } = null!;

    /// <inheritdoc />
    protected override void Execute(CodeActivityContext context)
    {
        var archive = ZipFile?.Get(context);
        if (string.IsNullOrWhiteSpace(archive))
            throw new ArgumentException("Zip File is required.", nameof(ZipFile));

        var sources = FolderFilesPath?.Get(context) ?? [];
        if (sources.Length == 0)
            throw new ArgumentException("Folder Files Path is required.", nameof(FolderFilesPath));

        var service = context.GetExtension<Zip.IArchiveService>() ?? Zip.ArchiveService.Instance;
        service.Compress(new Zip.CompressRequest
        {
            ArchiveFile = archive,
            Sources = sources,
            Password = Password?.Get(context) ?? string.Empty,
        });
    }
}
