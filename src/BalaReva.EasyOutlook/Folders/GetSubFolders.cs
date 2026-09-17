using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook.Folders;

/// <summary>Lists the folders nested under a named folder.</summary>
[DisplayName("Get Sub Folders")]
[Description("Lists the names of the folders nested under a named folder.")]
public sealed class GetSubFolders : BaseActivity
{
    /// <summary>Folder whose children to list.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Folder")]
    [Description("Name of the folder whose children to list.")]
    public InArgument<string> Folder { get; set; } = null!;

    /// <summary>Names of the nested folders.</summary>
    [Category("Output")]
    [DisplayName("Folders List")]
    [Description("Names of the folders nested under the named folder.")]
    public OutArgument<string[]> FoldersList { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service) =>
        FoldersList.Set(context, service.GetSubFolders(Require(context, Folder, nameof(Folder))));
}
