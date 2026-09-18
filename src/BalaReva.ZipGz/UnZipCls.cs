using System.Activities;
using System.ComponentModel;

namespace BalaReva.ZipGz;

/// <summary>Extracts a gzip or tar archive into a folder.</summary>
[DisplayName("UnZip Gz")]
[Description("Extracts a gzip or tar archive into a folder.")]
public sealed class UnZipCls : CodeActivity
{
    /// <summary>Full path of the archive to extract.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Zip File")]
    [Description("Full path of the .gz, .tgz or .tar archive to extract.")]
    public InArgument<string> ZipFile { get; set; } = null!;

    /// <summary>Folder to extract into. Created when it does not exist.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Extract Folder Path")]
    [Description("Folder to extract into. Created when it does not exist.")]
    public InArgument<string> ExtractFolderPath { get; set; } = null!;

    /// <inheritdoc />
    protected override void Execute(CodeActivityContext context)
    {
        var archive = ZipFile?.Get(context);
        if (string.IsNullOrWhiteSpace(archive))
            throw new ArgumentException("Zip File is required.", nameof(ZipFile));

        var destination = ExtractFolderPath?.Get(context);
        if (string.IsNullOrWhiteSpace(destination))
            throw new ArgumentException("Extract Folder Path is required.", nameof(ExtractFolderPath));

        var service = context.GetExtension<IGzipService>() ?? GzipService.Instance;
        service.Extract(archive, destination);
    }
}
