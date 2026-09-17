using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook.Folders;

/// <summary>Deletes a nested folder, and everything in it.</summary>
[DisplayName("Delete Sub Folder")]
[Description("Deletes a nested folder, along with everything inside it.")]
public sealed class DeleteSubFolder : BaseActivity
{
    /// <summary>Folder holding the one to delete.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Folder Name")]
    [Description("Name of the folder holding the one to delete.")]
    public InArgument<string> FolderName { get; set; } = null!;

    /// <summary>Nested folder to delete.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Sub Folder Name")]
    [Description("Name of the nested folder to delete.")]
    public InArgument<string> SubFolderName { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service) =>
        service.DeleteSubFolder(
            Require(context, FolderName, nameof(FolderName)),
            Require(context, SubFolderName, nameof(SubFolderName)));
}
