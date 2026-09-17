using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook.Folders;

/// <summary>Creates a folder under the mailbox root.</summary>
[DisplayName("Create Folder")]
[Description("Creates a folder directly under the mailbox root.")]
public sealed class CreateFolder : BaseActivity
{
    /// <summary>Name for the new folder.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Folder Name")]
    [Description("Name for the new folder.")]
    public InArgument<string> FolderName { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service) =>
        service.CreateFolder(Require(context, FolderName, nameof(FolderName)));
}
