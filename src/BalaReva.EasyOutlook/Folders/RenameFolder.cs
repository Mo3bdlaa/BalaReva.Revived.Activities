using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook.Folders;

/// <summary>Renames a folder under the mailbox root.</summary>
[DisplayName("Rename Folder")]
[Description("Renames a folder under the mailbox root.")]
public sealed class RenameFolder : BaseActivity
{
    /// <summary>Folder to rename.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Existing Folder Name")]
    [Description("Current name of the folder.")]
    public InArgument<string> ExistingFolderName { get; set; } = null!;

    /// <summary>New name.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("New Folder Name")]
    [Description("Name to give the folder.")]
    public InArgument<string> NewFolderName { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service) =>
        service.RenameFolder(
            Require(context, ExistingFolderName, nameof(ExistingFolderName)),
            Require(context, NewFolderName, nameof(NewFolderName)));
}
