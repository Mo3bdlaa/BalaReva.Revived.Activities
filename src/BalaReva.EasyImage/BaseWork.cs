using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyImage;

/// <summary>
/// Shared behaviour for every image activity: the optional delay, the error
/// swallowing, and reporting whether the step ran.
/// </summary>
public abstract class BaseWork : CodeActivity
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
    public InArgument<short> Delay { get; set; } = null!;

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
            ExecuteWork(context);
            ExecutionResult.Set(context, true);
        }
        catch (Exception) when (ContinueOnError.Get(context))
        {
            ExecutionResult.Set(context, false);
        }
    }

    /// <summary>Does the actual work.</summary>
    protected abstract void ExecuteWork(CodeActivityContext context);

    /// <summary>Reads a required path argument and checks the file is there.</summary>
    protected static string RequireExistingFile(
        CodeActivityContext context, InArgument<string> argument, string name)
    {
        var path = RequirePath(context, argument, name);
        if (!File.Exists(path))
            throw new FileNotFoundException($"The file '{path}' does not exist.", path);
        return path;
    }

    /// <summary>Reads a required path argument without touching the file system.</summary>
    protected static string RequirePath(
        CodeActivityContext context, InArgument<string> argument, string name)
    {
        var path = argument?.Get(context);
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException($"{name} is required.", name);
        return path;
    }
}
