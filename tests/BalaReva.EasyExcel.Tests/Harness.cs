using System.Activities;
using System.Activities.Statements;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Main;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.Tests;

/// <summary>Whatever the captured output arguments held when the workflow finished.</summary>
internal sealed class Outputs
{
    public Dictionary<string, object> Values { get; } = [];
}

/// <summary>Copies a workflow variable into <see cref="Outputs"/> once the scope is done.</summary>
internal sealed class CaptureOutput<T>(Outputs sink, string name) : CodeActivity
{
    public InArgument<T> Value { get; set; } = null!;

    protected override void Execute(CodeActivityContext context) =>
        sink.Values[name] = context.GetValue(Value)!;
}

/// <summary>
/// Builds a workflow that runs one activity inside an ExcelScope and reads its outputs
/// back, which a nested activity cannot return on its own.
/// </summary>
internal sealed class ScopeRun
{
    private readonly List<Variable> _variables = [];
    private readonly List<Activity> _captures = [];
    private readonly Outputs _outputs = new();

    public OutArgument<T> Capture<T>(string name)
    {
        var variable = new Variable<T>();
        _variables.Add(variable);
        _captures.Add(new CaptureOutput<T>(_outputs, name) { Value = new InArgument<T>(variable) });
        return new OutArgument<T>(variable);
    }

    public Outputs Invoke(ExcelActivity child, FakeExcelService service,
                          string filePath = @"C:\books\book.xlsx",
                          EnableDisableEnum macros = EnableDisableEnum.Disable)
    {
        child.ContinueOnError ??= new InArgument<bool>(false);
        child.Delay ??= new InArgument<short>(0);
        child.ExecutionResult ??= new OutArgument<bool>();

        var scope = new ExcelScope
        {
            FilePath = new InArgument<string>(filePath),
            MacroSettings = macros,
            Body = new ActivityAction<ExcelParam>
            {
                Argument = new DelegateInArgument<ExcelParam> { Name = "ExcelWorkBook" },
                Handler = child,
            },
        };

        var root = new Sequence { Activities = { scope } };
        foreach (var variable in _variables) root.Variables.Add(variable);
        foreach (var capture in _captures) root.Activities.Add(capture);

        var invoker = new WorkflowInvoker(root);
        invoker.Extensions.Add(service);
        invoker.Invoke();
        return _outputs;
    }
}

/// <summary>Runs one in-scope activity and returns the captured outputs.</summary>
internal static class Harness
{
    public static IDictionary<string, object> Run(
        ExcelActivity child, FakeExcelService service,
        string filePath = @"C:\books\book.xlsx") =>
        new ScopeRun().Invoke(child, service, filePath).Values;

    /// <summary>Runs an activity that stands outside any scope.</summary>
    public static void RunAlone(Activity activity, FakeExcelService service)
    {
        var invoker = new WorkflowInvoker(activity);
        invoker.Extensions.Add(service);
        invoker.Invoke();
    }
}
