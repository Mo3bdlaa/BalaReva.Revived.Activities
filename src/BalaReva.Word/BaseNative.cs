using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word;

/// <summary>Shared arguments for the activity that opens a document.</summary>
public abstract class BaseNative : NativeActivity
{
    /// <summary>Full path of the document.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("File Path")]
    [Description("Full path of the Word document to open.")]
    public InArgument<string> FilePath { get; set; } = null!;

    /// <summary>Password needed to open the document.</summary>
    [Category("Input")]
    [DisplayName("Open Password")]
    [Description("Password needed to open the document, if it has one.")]
    public InArgument<string> OpenPassword { get; set; } = null!;

    /// <summary>Password needed to change the document.</summary>
    [Category("Input")]
    [DisplayName("Modify Password")]
    [Description("Password needed to change the document, if it has one.")]
    public InArgument<string> ModifyPassword { get; set; } = null!;

    /// <summary>
    /// The Word session to talk to: a workflow extension when one is registered,
    /// otherwise the real COM implementation.
    /// </summary>
    private protected static IWordService Service(NativeActivityContext context) =>
        context.GetExtension<IWordService>() ?? WordService.Instance;
}
