using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook.Folders;

/// <summary>Creates a folder nested under an existing one.</summary>
[DisplayName("Create Sub Folder")]
[Description("Creates a folder nested under an existing folder.")]
public sealed class CreateSubFolder : BaseActivity
{
    /// <summary>Folder to create the new one inside.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Folder")]
    [Description("Name of the folder to create the new one inside.")]
    public InArgument<string> Folder { get; set; } = null!;

    /// <summary>Name for the new nested folder.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Sub Folder")]
    [Description("Name for the new nested folder.")]
    public InArgument<string> SubFolder { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service) =>
        service.CreateSubFolder(
            Require(context, Folder, nameof(Folder)),
            Require(context, SubFolder, nameof(SubFolder)));
}
