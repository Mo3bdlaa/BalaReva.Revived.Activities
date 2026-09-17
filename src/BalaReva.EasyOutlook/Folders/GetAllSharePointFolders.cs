using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook.Folders;

/// <summary>Lists the SharePoint list folders connected to Outlook.</summary>
[DisplayName("Get All SharePoint Folders")]
[Description("Lists the names of the SharePoint list folders connected to Outlook.")]
public sealed class GetAllSharePointFolders : BaseActivity
{
    /// <summary>Names of the connected SharePoint list folders.</summary>
    [Category("Output")]
    [DisplayName("Folders List")]
    [Description("Names of the connected SharePoint list folders.")]
    public OutArgument<string[]> FoldersList { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service) =>
        FoldersList.Set(context, service.GetSharePointFolders());
}
