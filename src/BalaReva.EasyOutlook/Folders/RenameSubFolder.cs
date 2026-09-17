using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook.Folders;

/// <summary>Renames a folder nested under another.</summary>
[DisplayName("Rename Sub Folder")]
[Description("Renames a folder nested under another folder.")]
public sealed class RenameSubFolder : BaseActivity
{
    /// <summary>Folder holding the one to rename.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Folder Name")]
    [Description("Name of the folder holding the one to rename.")]
    public InArgument<string> FolderName { get; set; } = null!;

    /// <summary>Nested folder to rename.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Existing Sub Folder Name")]
    [Description("Current name of the nested folder.")]
    public InArgument<string> ExistingSubFolderName { get; set; } = null!;

    /// <summary>New name.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("New Sub Folder Name")]
    [Description("Name to give the nested folder.")]
    public InArgument<string> NewSubFolderName { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service) =>
        service.RenameSubFolder(
            Require(context, FolderName, nameof(FolderName)),
            Require(context, ExistingSubFolderName, nameof(ExistingSubFolderName)),
            Require(context, NewSubFolderName, nameof(NewSubFolderName)));
}
