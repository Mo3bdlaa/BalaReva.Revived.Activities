using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook.Folders;

/// <summary>Lists the folders under a connected SharePoint list folder.</summary>
[DisplayName("Get All SharePoint Sub Folders")]
[Description("Lists the names of the folders under a connected SharePoint list folder.")]
public sealed class GetAllSharePointSubFolders : BaseActivity
{
    /// <summary>SharePoint folder whose children to list.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Folder")]
    [Description("Name of the SharePoint list folder whose children to list.")]
    public InArgument<string> Folder { get; set; } = null!;

    /// <summary>Names of the nested folders.</summary>
    [Category("Output")]
    [DisplayName("Folders List")]
    [Description("Names of the folders under the named SharePoint list folder.")]
    public OutArgument<string[]> FoldersList { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service) =>
        FoldersList.Set(context, service.GetSharePointSubFolders(Require(context, Folder, nameof(Folder))));
}
