using System.Activities;

namespace BalaReva.Printer.Tests;

/// <summary>Runs one activity as a workflow root with a stand-in spooler registered.</summary>
internal static class Harness
{
    /// <summary>
    /// Invokes <paramref name="activity"/> and returns its output arguments by name.
    /// </summary>
    public static IDictionary<string, object> Run(BaseActivity activity, FakePrinterService service)
    {
        activity.ContinueOnError ??= new InArgument<bool>(false);
        activity.Delay ??= new InArgument<short>(0);
        activity.ExecutionResult ??= new OutArgument<bool>();

        var invoker = new WorkflowInvoker(activity);
        invoker.Extensions.Add(service);
        return invoker.Invoke();
    }
}
