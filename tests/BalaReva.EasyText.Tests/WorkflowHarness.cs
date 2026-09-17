using System.Activities;
using System.Activities.Statements;
using BalaReva.EasyText.Base;
using BalaReva.EasyText.Main;

namespace BalaReva.EasyText.Tests;

/// <summary>
/// Runs a single child activity inside a real <see cref="TextScope"/> on a real
/// temporary file, through the actual workflow runtime.
/// </summary>
/// <remarks>
/// These activities only work when the scope publishes its handle and the child
/// resolves it, so exercising them through <see cref="WorkflowInvoker"/> is the
/// only way to cover that contract. Calling Execute directly would skip it.
/// </remarks>
public sealed class ScopedFile : IDisposable
{
    public ScopedFile(string contents)
    {
        Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"easytext-{Guid.NewGuid():N}.txt");
        File.WriteAllText(Path, contents);
    }

    public string Path { get; }

    public string Contents => File.ReadAllText(Path);

    public byte[] Bytes => File.ReadAllBytes(Path);

    /// <summary>Runs <paramref name="child"/> inside a scope over this file.</summary>
    public void Run(BaseChildActivity child) => WorkflowInvoker.Invoke(Wrap(child, null));

    /// <summary>
    /// Runs <paramref name="child"/> inside a scope and returns whatever
    /// <paramref name="bind"/> routed into a variable.
    /// </summary>
    public T Run<T>(BaseChildActivity child, Func<BaseChildActivity, Variable<T>, Activity> bind)
    {
        var captured = new List<T>();
        var variable = new Variable<T>();
        var wired = bind(child, variable);
        WorkflowInvoker.Invoke(new Sequence
        {
            Variables = { variable },
            Activities = { Wrap(wired, null), new Capture<T>(captured) { Value = variable } },
        });
        return captured.Single();
    }

    private Activity Wrap(Activity child, DelegateInArgument<string>? argument) => new TextScope
    {
        FilePath = new InArgument<string>(Path),
        Body = new ActivityAction<string>
        {
            Argument = argument ?? new DelegateInArgument<string> { Name = "FilePath" },
            Handler = child,
        },
    };

    public void Dispose()
    {
        if (File.Exists(Path)) File.Delete(Path);
    }
}

/// <summary>Copies a workflow variable out into a list the test can assert on.</summary>
internal sealed class Capture<T>(List<T> sink) : CodeActivity
{
    public InArgument<T> Value { get; set; } = null!;

    protected override void Execute(CodeActivityContext context) => sink.Add(Value.Get(context));
}
