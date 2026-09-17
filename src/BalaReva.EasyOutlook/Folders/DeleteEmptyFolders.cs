using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook.Folders;

/// <summary>Deletes every empty folder under the mailbox root.</summary>
/// <remarks>
/// A folder counts as empty when it holds no items and no nested folders. The sweep is
/// one level deep, so a folder whose only child is itself empty is left alone.
/// </remarks>
[DisplayName("Delete Empty Folders")]
[Description("Deletes every folder under the mailbox root that holds no items and no sub folders.")]
public sealed class DeleteEmptyFolders : BaseActivity
{
    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service) =>
        service.DeleteEmptyFolders();
}
