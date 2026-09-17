using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyText.Base;

/// <summary>
/// Shared behaviour for every activity that runs inside a <c>TextScope</c>:
/// the optional delay, the error swallowing, and reporting whether the step ran.
/// </summary>
public abstract class BaseChildActivity : CodeActivity
{
    /// <summary>
    /// When true, a failure is reported through <see cref="ExecutionResult"/>
    /// instead of faulting the workflow.
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
    public InArgument<int> Delay { get; set; } = null!;

    /// <summary>True when the activity completed without error.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the activity completed without error.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected sealed override void Execute(CodeActivityContext context)
    {
        var delay = Delay.Get(context);
        if (delay > 0) Thread.Sleep(delay);

        try
        {
            ExecuteChild(context, ParentActivityValidation.Resolve(context));
            ExecutionResult.Set(context, true);
        }
        catch (Exception) when (ContinueOnError.Get(context))
        {
            ExecutionResult.Set(context, false);
        }
    }

    /// <summary>Does the actual work against the scope's <paramref name="document"/>.</summary>
    protected abstract void ExecuteChild(CodeActivityContext context, TextDocument document);

    /// <summary>
    /// Validates a 1-based line number against a file of <paramref name="lineCount"/> lines
    /// and returns the matching 0-based index.
    /// </summary>
    /// <remarks>
    /// Line numbers on these activities are 1-based, matching how line numbers are
    /// shown in editors. Character offsets, such as FindText's, stay 0-based to match
    /// <see cref="string.IndexOf(string, StringComparison)"/>.
    /// </remarks>
    protected static int LineIndex(int lineNumber, int lineCount, string argumentName)
    {
        if (lineNumber < 1 || lineNumber > lineCount)
        {
            throw new ArgumentOutOfRangeException(
                argumentName, lineNumber,
                $"Line number must be between 1 and {lineCount}.");
        }
        return lineNumber - 1;
    }
}
