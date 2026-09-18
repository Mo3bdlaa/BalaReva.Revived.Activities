using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Individual.Presentation;

/// <summary>Runs a PowerPoint VBA macro.</summary>
/// <remarks>
/// Opens the presentation itself rather than running inside a scope, which is how the
/// published package had it, so it carries its own file path and passwords.
/// </remarks>
[DisplayName("Execute Macro")]
[Description("Runs a PowerPoint VBA macro by name and returns its result.")]
public sealed class ExecuteMacro : BasePowerPoint
{
    /// <summary>Name of the macro.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Macro Name")]
    [Description("Name of the macro to run, for example Module1.MyMacro.")]
    public InArgument<string> MacroName { get; set; } = null!;

    /// <summary>Arguments to pass to the macro.</summary>
    [Category("Input")]
    [DisplayName("Arguments")]
    [Description("Arguments to pass to the macro, in order.")]
    public InArgument<object[]> Arguments { get; set; } = null!;

    /// <summary>Whatever the macro returned.</summary>
    [Category("Output")]
    [DisplayName("Result")]
    [Description("The macro's return value, or null when it returns nothing.")]
    public OutArgument<object> Result { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
    {
        var name = MacroName?.Get(context);
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("MacroName is required.", nameof(MacroName));

        // Null is fine in the argument even though the signature is not annotated for it.
        Result.Set(context, presentation.ExecuteMacro(name, Arguments?.Get(context))!);
    }
}
