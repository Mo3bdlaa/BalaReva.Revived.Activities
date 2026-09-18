using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word.Documents;

/// <summary>Sets new open and modify passwords on the document.</summary>
/// <remarks>
/// The name is wrong, and it is wrong in the published package too: this activity's only
/// arguments are NewOpenPassword and NewModifyPassword, which is a password change, not
/// a find and replace. It looks like a copy of <see cref="ChangePassword"/> that was
/// renamed and shipped. Renaming it here would break any workflow that binds to it, so
/// it keeps its name and does what its arguments say. For find and replace, use
/// <c>BalaReva.Word.Pages.FindReplace</c>, which has the arguments for it.
/// </remarks>
[DisplayName("Find And Replace")]
[Description("Sets new open and modify passwords. Despite the name, this is a password "
             + "change - see Pages.FindReplace for find and replace.")]
public sealed class FindAndReplace : BaseNativeChild
{
    /// <summary>New password for opening the document.</summary>
    [Category("Input")]
    [DisplayName("New Open Password")]
    [Description("New password for opening the document. Empty removes it.")]
    public InArgument<string> NewOpenPassword { get; set; } = null!;

    /// <summary>New password for changing the document.</summary>
    [Category("Input")]
    [DisplayName("New Modify Password")]
    [Description("New password for changing the document. Empty removes it.")]
    public InArgument<string> NewModifyPassword { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.ChangePassword(
            NewOpenPassword?.Get(context) ?? string.Empty,
            NewModifyPassword?.Get(context) ?? string.Empty);
}
