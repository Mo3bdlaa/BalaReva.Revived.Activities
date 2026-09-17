using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook.Folders;

/// <summary>Deletes a folder under the mailbox root, and everything in it.</summary>
[DisplayName("Delete Folder")]
[Description("Deletes a folder under the mailbox root, along with everything inside it.")]
public sealed class DeleteFolder : BaseActivity
{
    /// <summary>Folder to delete.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Folder Name")]
    [Description("Name of the folder to delete.")]
    public InArgument<string> FolderName { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service) =>
        service.DeleteFolder(Require(context, FolderName, nameof(FolderName)));
}
