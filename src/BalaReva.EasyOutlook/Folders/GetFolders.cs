using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook.Folders;

/// <summary>Lists the folders directly under the mailbox root.</summary>
[DisplayName("Get Folders")]
[Description("Lists the names of the folders directly under the mailbox root.")]
public sealed class GetFolders : BaseActivity
{
    /// <summary>Names of the top-level folders.</summary>
    [Category("Output")]
    [DisplayName("Folders List")]
    [Description("Names of the folders directly under the mailbox root.")]
    public OutArgument<string[]> FoldersList { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service) =>
        FoldersList.Set(context, service.GetFolders());
}
