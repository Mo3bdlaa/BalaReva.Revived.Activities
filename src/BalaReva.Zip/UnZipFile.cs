using System.Activities;
using System.ComponentModel;

namespace BalaReva.ZipUnzip;

/// <summary>Extracts an archive into a folder.</summary>
/// <remarks>
/// Note <c>strZipFile</c>: that is how the published package spelled it, and a workflow
/// binds by property name, so the Hungarian prefix stays.
/// </remarks>
[DisplayName("UnZip File")]
[Description("Extracts an archive into a folder.")]
public sealed class UnZipFile : CodeActivity
{
    /// <summary>Full path of the archive to extract.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Zip File")]
    [Description("Full path of the archive to extract.")]
    public InArgument<string> strZipFile { get; set; } = null!;

    /// <summary>Folder to extract into. Created when it does not exist.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Extract Folder Path")]
    [Description("Folder to extract into. Created when it does not exist.")]
    public InArgument<string> ExtractFolderPath { get; set; } = null!;

    /// <summary>Password, when the archive needs one.</summary>
    [Category("Input")]
    [DisplayName("Password")]
    [Description("Password, when the archive needs one.")]
    public InArgument<string> Password { get; set; } = null!;

    /// <summary>Zip only, or any format the readers recognise.</summary>
    [Category("Input")]
    [DisplayName("Extract Type")]
    [Description("Standard reads zip archives; UniExtract reads rar, 7z, tar and the rest.")]
    public Zip.EnumExtractType ExtractType { get; set; } = Zip.EnumExtractType.Standard;

    /// <summary>
    /// Code page for entry names, for archives written by a tool that did not use UTF-8.
    /// Zero leaves the default alone.
    /// </summary>
    [Category("Input")]
    [DisplayName("Code Page")]
    [Description("Code page for entry names. Zero leaves the default alone.")]
    public InArgument<int> CodePage { get; set; } = null!;

    /// <inheritdoc />
    protected override void Execute(CodeActivityContext context)
    {
        var archive = strZipFile?.Get(context);
        if (string.IsNullOrWhiteSpace(archive))
            throw new ArgumentException("Zip File is required.", nameof(strZipFile));

        var destination = ExtractFolderPath?.Get(context);
        if (string.IsNullOrWhiteSpace(destination))
            throw new ArgumentException("Extract Folder Path is required.", nameof(ExtractFolderPath));

        var service = context.GetExtension<Zip.IArchiveService>() ?? Zip.ArchiveService.Instance;
        service.Extract(new Zip.ExtractRequest
        {
            ArchiveFile = archive,
            DestinationFolder = destination,
            Password = Password?.Get(context) ?? string.Empty,
            ExtractType = ExtractType,
            CodePage = CodePage?.Get(context) ?? 0,
        });
    }
}
