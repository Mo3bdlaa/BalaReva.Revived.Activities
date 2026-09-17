using System.Activities;
using System.ComponentModel;

namespace BalaReva.EasyOutlook;

/// <summary>
/// Shared behaviour for every Outlook activity.
/// </summary>
/// <remarks>
/// Only a delay, deliberately. Unlike the EasyImage and Printer packages, the published
/// EasyOutlook base carries no ContinueOnError or ExecutionResult, so failures here
/// fault the workflow. Adding them would change the property grid every one of these
/// activities presents, so the shape is kept as published.
/// </remarks>
public abstract class BaseActivity : CodeActivity
{
    /// <summary>Milliseconds to wait before running.</summary>
    [RequiredArgument]
    [Category("Common")]
    [DisplayName("Delay")]
    [Description("Milliseconds to wait before this activity runs.")]
    public InArgument<short> Delay { get; set; } = null!;

    /// <inheritdoc />
    protected sealed override void Execute(CodeActivityContext context)
    {
        var delay = Delay.Get(context);
        if (delay > 0) Thread.Sleep(delay);
        ExecuteWork(context, Service(context));
    }

    /// <summary>Does the actual work against Outlook.</summary>
    protected abstract void ExecuteWork(CodeActivityContext context, IOutlookService service);

    /// <summary>
    /// The Outlook session to talk to: a workflow extension when one is registered,
    /// otherwise the real COM implementation.
    /// </summary>
    private static IOutlookService Service(CodeActivityContext context) =>
        context.GetExtension<IOutlookService>() ?? OutlookService.Instance;

    /// <summary>Reads a required string argument.</summary>
    protected static string Require(CodeActivityContext context, InArgument<string> argument, string name)
    {
        var value = argument?.Get(context);
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException($"{name} is required.", name);
        return value;
    }
}
