using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook.Contacts;

/// <summary>Adds a contact to Outlook.</summary>
[DisplayName("New Contact")]
[Description("Adds a contact to the Contacts folder.")]
public sealed class NewContact : BaseActivity
{
    /// <summary>Outlook account to add the contact to. Leave empty for the default.</summary>
    [Category("Input")]
    [DisplayName("Account")]
    [Description("Outlook account to add the contact to. Leave empty for the default account.")]
    public InArgument<string> Account { get; set; } = null!;

    /// <summary>The contact's details.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Contact")]
    [Description("The contact's details. Anything left unset is left unset on the contact.")]
    public InArgument<OutlookNewContact> Contact { get; set; } = null!;

    /// <summary>True when the contact was saved.</summary>
    [Category("Output")]
    [DisplayName("Added Result")]
    [Description("True when the contact was saved.")]
    public OutArgument<bool> AddedResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service)
    {
        var contact = Contact?.Get(context)
            ?? throw new ArgumentException("Contact is required.", nameof(Contact));

        AddedResult.Set(context, service.AddContact(Account?.Get(context) ?? string.Empty, contact));
    }
}
