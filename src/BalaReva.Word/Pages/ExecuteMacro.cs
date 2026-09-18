using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word.Pages;

/// <summary>Runs a Word VBA macro.</summary>
[DisplayName("Execute Macro")]
[Description("Runs a Word VBA macro by name and returns its result.")]
public sealed class ExecuteMacro : BaseNativeChild
{
    /// <summary>Name of the macro.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Macro Name")]
    [Description("Name of the macro to run, for example Project.Module.MyMacro.")]
    public InArgument<string> MacroName { get; set; } = null!;

    /// <summary>Arguments to pass to the macro.</summary>
    [Category("Input")]
    [DisplayName("Parameters")]
    [Description("Arguments to pass to the macro, in order.")]
    public InArgument<object[]> Parameters { get; set; } = null!;

    /// <summary>Whatever the macro returned.</summary>
    [Category("Output")]
    [DisplayName("Macro Output")]
    [Description("The macro's return value, or null when it returns nothing.")]
    public OutArgument<object> MacroOutput { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document)
    {
        var result = document.ExecuteMacro(
            Require(context, MacroName, nameof(MacroName)),
            Parameters?.Get(context));

        // Null is fine in the argument even though the signature is not annotated for it.
        MacroOutput.Set(context, result!);
    }
}
