using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook.Folders;

/// <summary>Deletes every empty folder nested under a named folder.</summary>
/// <remarks>See the note on <see cref="DeleteEmptyFolders"/> about what counts as empty.</remarks>
[DisplayName("Delete Empty Sub Folders")]
[Description("Deletes every folder nested under the named folder that holds no items and no sub folders.")]
public sealed class DeleteEmptySubFolders : BaseActivity
{
    /// <summary>Folder to sweep.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Folder Name")]
    [Description("Name of the folder whose empty children to delete.")]
    public InArgument<string> FolderName { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service) =>
        service.DeleteEmptySubFolders(Require(context, FolderName, nameof(FolderName)));
}
