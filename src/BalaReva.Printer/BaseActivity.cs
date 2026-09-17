using System.Activities;
using System.ComponentModel;

namespace BalaReva.Printer;

/// <summary>
/// Shared behaviour for every printer activity: the optional delay, the error
/// swallowing, and reporting whether the step ran.
/// </summary>
public abstract class BaseActivity : CodeActivity
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
            ExecuteWork(context, Service(context));
            ExecutionResult.Set(context, true);
        }
        catch (Exception) when (ContinueOnError.Get(context))
        {
            ExecutionResult.Set(context, false);
        }
    }

    /// <summary>Does the actual work against the spooler.</summary>
    protected abstract void ExecuteWork(CodeActivityContext context, IPrinterService service);

    /// <summary>
    /// The spooler to talk to: a workflow extension when one is registered, otherwise
    /// the real Windows implementation.
    /// </summary>
    private static IPrinterService Service(CodeActivityContext context) =>
        context.GetExtension<IPrinterService>() ?? WindowsPrinterService.Instance;

    /// <summary>Reads a required printer name argument.</summary>
    protected static string RequirePrinterName(CodeActivityContext context, InArgument<string> argument)
    {
        var name = argument?.Get(context);
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("PrinterName is required.", nameof(argument));
        return name;
    }
}
