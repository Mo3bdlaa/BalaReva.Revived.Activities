using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook.Notes;

/// <summary>Adds a note to Outlook.</summary>
[DisplayName("New Notes")]
[Description("Adds a note to the default Notes folder.")]
public sealed class NewNotes : BaseActivity
{
    /// <summary>Text of the note.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Note Body")]
    [Description("Text of the note. Outlook takes the first line as its title.")]
    public InArgument<string> NoteBody { get; set; } = null!;

    /// <summary>True when the note was saved.</summary>
    [Category("Output")]
    [DisplayName("Added Result")]
    [Description("True when the note was saved.")]
    public OutArgument<bool> AddedResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service) =>
        AddedResult.Set(context, service.AddNote(Require(context, NoteBody, nameof(NoteBody))));
}
