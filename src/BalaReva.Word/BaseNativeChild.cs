using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word;

/// <summary>
/// Shared behaviour for every activity that runs inside a <c>WordScope</c>.
/// </summary>
/// <remarks>
/// Note that <c>Delay</c> is a double here, where the other revived packages use a short
/// or an int. That is how the published package declared it, and a workflow binds by
/// type as well as by name.
/// </remarks>
public abstract class BaseNativeChild : CodeActivity
{
    /// <summary>
    /// When true, a failure is reported through <see cref="ExecutionResult"/> instead of
    /// faulting the workflow.
    /// </summary>
    [RequiredArgument]
    [Category("Common")]
    [DisplayName("Continue On Error")]
    [Description("Report failures through ExecutionResult instead of faulting the workflow.")]
    public InArgument<bool> ContinueOnError { get; set; } = null!;

    /// <summary>Milliseconds to wait before running.</summary>
    [RequiredArgument]
    [Category("Common")]
    [DisplayName("Delay")]
    [Description("Milliseconds to wait before this activity runs.")]
    public InArgument<double> Delay { get; set; } = null!;

    /// <summary>True when the activity completed without error.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the activity completed without error.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected sealed override void Execute(CodeActivityContext context)
    {
        var delay = Delay.Get(context);
        if (delay > 0) Thread.Sleep(TimeSpan.FromMilliseconds(delay));

        try
        {
            ExecuteWork(context, Resolve(context));
            ExecutionResult.Set(context, true);
        }
        catch (Exception) when (ContinueOnError.Get(context))
        {
            ExecutionResult.Set(context, false);
        }
    }

    /// <summary>Does the actual work against the scope's open document.</summary>
    protected abstract void ExecuteWork(CodeActivityContext context, IWordDocument document);

    /// <summary>Finds the document opened by the enclosing scope.</summary>
    private static IWordDocument Resolve(CodeActivityContext context) =>
        context.GetProperty<WordScopeHandle>()?.Document
        ?? throw new InvalidOperationException(
            "This activity must be placed inside a Word Scope activity, which opens the document.");

    /// <summary>Reads a required string argument.</summary>
    protected static string Require(CodeActivityContext context, InArgument<string> argument, string name)
    {
        var value = argument?.Get(context);
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException($"{name} is required.", name);
        return value;
    }
}
