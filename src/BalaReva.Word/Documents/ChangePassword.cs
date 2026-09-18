using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word.Documents;

/// <summary>Sets new open and modify passwords on the document.</summary>
[DisplayName("Change Password")]
[Description("Sets new open and modify passwords on the document.")]
public sealed class ChangePassword : BaseNativeChild
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
