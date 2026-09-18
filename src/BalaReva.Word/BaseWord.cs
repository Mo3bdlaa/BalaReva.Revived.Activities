using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word;

/// <summary>
/// Shared behaviour for the activities that work on a file directly, without a scope.
/// </summary>
/// <remarks>
/// The published package gives this base no Delay and no ExecutionResult, unlike
/// <see cref="BaseNativeChild"/>, so its activities fault the workflow on failure unless
/// ContinueOnError is set.
/// </remarks>
public abstract class BaseWord : CodeActivity
{
    /// <summary>When true, failures are swallowed rather than faulting the workflow.</summary>
    [RequiredArgument]
    [Category("Common")]
    [DisplayName("Continue On Error")]
    [Description("Swallow failures instead of faulting the workflow.")]
    public InArgument<bool> ContinueOnError { get; set; } = null!;

    /// <summary>Full path of the document.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Word File")]
    [Description("Full path of the Word document.")]
    public InArgument<string> WordFile { get; set; } = null!;

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

    /// <inheritdoc />
    protected sealed override void Execute(CodeActivityContext context)
    {
        try
        {
            ExecuteWork(context, context.GetExtension<IWordService>() ?? WordService.Instance);
        }
        catch (Exception) when (ContinueOnError.Get(context))
        {
            // Reported nowhere: this base has no ExecutionResult to set.
        }
    }

    /// <summary>Does the actual work.</summary>
    protected abstract void ExecuteWork(CodeActivityContext context, IWordService service);

    /// <summary>Reads a required string argument.</summary>
    protected static string Require(CodeActivityContext context, InArgument<string> argument, string name)
    {
        var value = argument?.Get(context);
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException($"{name} is required.", name);
        return value;
    }
}
