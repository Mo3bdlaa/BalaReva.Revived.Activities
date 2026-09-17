using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook.Notes;

/// <summary>Reads the body text of every Outlook note.</summary>
[DisplayName("Get All Notes")]
[Description("Reads the body text of every note in the default Notes folder.")]
public sealed class GetAllNotes : BaseActivity
{
    /// <summary>Body text of each note.</summary>
    [Category("Output")]
    [DisplayName("Notes Collection")]
    [Description("Body text of each note in the default Notes folder.")]
    public OutArgument<List<string>> NotesCollection { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service) =>
        NotesCollection.Set(context, service.GetNotes());
}
