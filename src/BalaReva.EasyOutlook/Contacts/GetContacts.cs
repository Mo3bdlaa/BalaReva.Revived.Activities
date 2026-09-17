using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook.Contacts;

/// <summary>Reads every contact out of Outlook.</summary>
[DisplayName("Get Contacts")]
[Description("Reads every contact in the default Contacts folder.")]
public sealed class GetContacts : BaseActivity
{
    /// <summary>The contacts, as snapshots.</summary>
    [Category("Output")]
    [DisplayName("Contact Collection")]
    [Description("Every contact in the default Contacts folder.")]
    public OutArgument<List<OutlookContact>> ContactCollection { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service) =>
        ContactCollection.Set(context, service.GetContacts());
}
