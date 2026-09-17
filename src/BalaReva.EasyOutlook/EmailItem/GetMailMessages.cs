using System.Activities;
using System.ComponentModel;
using BalaReva.EasyOutlook.Utilities;

namespace BalaReva.EasyOutlook.EmailItem;

/// <summary>Reads messages out of an Outlook mail folder.</summary>
[DisplayName("Get Mail Messages")]
[Description("Reads messages out of an Outlook mail folder.")]
public sealed class GetMailMessages : BaseActivity
{
    /// <summary>Outlook account to read from. Leave empty for the default.</summary>
    [Category("Input")]
    [DisplayName("Account")]
    [Description("Outlook account to read from. Leave empty for the default account.")]
    public InArgument<string> Account { get; set; } = null!;

    /// <summary>Which default folder to read.</summary>
    [Category("Input")]
    [DisplayName("Mail Folder")]
    [Description("Which default folder to read. Ignored when Specific Folder is set.")]
    public MailFolderEnum MailFolder { get; set; } = MailFolderEnum.Inbox;

    /// <summary>A named folder to read instead of a default one.</summary>
    [Category("Input")]
    [DisplayName("Specific Folder")]
    [Description("Name of a folder under the mailbox root to read instead of a default folder.")]
    public InArgument<string> SpecificFolder { get; set; } = null!;

    /// <summary>An Outlook restriction to narrow the messages.</summary>
    [Category("Input")]
    [DisplayName("Filter")]
    [Description("An Outlook restriction, in DASL or Jet syntax, applied to the folder's items.")]
    public InArgument<string> Filter { get; set; } = null!;

    /// <summary>The messages that matched.</summary>
    [Category("Output")]
    [DisplayName("Mail Items")]
    [Description("The messages read from the folder.")]
    public OutArgument<global::BalaReva.Outlook.EmailItem[]> MailItems { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service) =>
        MailItems.Set(context, service.GetMailMessages(
            Account?.Get(context) ?? string.Empty,
            MailFolder,
            SpecificFolder?.Get(context) ?? string.Empty,
            Filter?.Get(context) ?? string.Empty));
}
