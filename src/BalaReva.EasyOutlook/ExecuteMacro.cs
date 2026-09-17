using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook;

/// <summary>Runs an Outlook VBA macro.</summary>
/// <remarks>
/// Outlook's object model, unlike Excel's and Word's, has no documented way to run a
/// macro by name, so this goes through late binding. Whether it works depends on the
/// Outlook build and its macro security settings. See docs/REVIVAL.md.
/// </remarks>
[DisplayName("Execute Macro")]
[Description("Runs an Outlook VBA macro by name and returns its result.")]
public sealed class ExecuteMacro : BaseActivity
{
    /// <summary>Name of the macro.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Macro Name")]
    [Description("Name of the macro to run, for example Project1.Module1.MyMacro.")]
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
    protected override void ExecuteWork(CodeActivityContext context, IOutlookService service)
    {
        var result = service.ExecuteMacro(
            Require(context, MacroName, nameof(MacroName)),
            Arguments?.Get(context));

        // A macro that returns nothing yields null, which the argument holds happily
        // even though its signature is not annotated for it.
        Result.Set(context, result!);
    }
}
