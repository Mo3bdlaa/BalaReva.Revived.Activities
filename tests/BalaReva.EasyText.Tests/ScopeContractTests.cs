using System.Activities;
using BalaReva.EasyText.Base;
using BalaReva.EasyText.TextFile;

namespace BalaReva.EasyText.Tests;

public class ScopeContractTests
{
    [Fact]
    public void A_child_outside_a_scope_says_so()
    {
        var activity = new LineCount
        {
            ContinueOnError = new InArgument<bool>(false),
            Delay = new InArgument<int>(0),
            Result = new OutArgument<int>(),
        };

        var error = Assert.ThrowsAny<Exception>(() => WorkflowInvoker.Invoke(activity));
        var message = Unwrap(error).Message;
        Assert.Contains("Text Scope", message, StringComparison.Ordinal);
    }

    [Fact]
    public void ContinueOnError_reports_failure_instead_of_faulting()
    {
        using var file = new ScopedFile("alpha\nbravo\n");
        var activity = new DeleteAt
        {
            ContinueOnError = new InArgument<bool>(true),
            Delay = new InArgument<int>(0),
            // Out of range: this would otherwise fault the workflow.
            DeleteLineIndex = new InArgument<int>(99),
        };

        var executed = file.Run<bool>(activity, (a, v) =>
        {
            ((DeleteAt)a).ExecutionResult = new OutArgument<bool>(v);
            return a;
        });

        Assert.False(executed);
        Assert.Equal("alpha\nbravo\n", file.Contents);
    }

    [Fact]
    public void Without_ContinueOnError_a_failure_faults_the_workflow()
    {
        using var file = new ScopedFile("alpha\nbravo\n");
        var activity = new DeleteAt
        {
            ContinueOnError = new InArgument<bool>(false),
            Delay = new InArgument<int>(0),
            DeleteLineIndex = new InArgument<int>(99),
        };

        Assert.ThrowsAny<Exception>(() => file.Run(activity));
        Assert.Equal("alpha\nbravo\n", file.Contents);
    }

    [Fact]
    public void A_successful_activity_reports_ExecutionResult_true()
    {
        using var file = new ScopedFile("alpha\nbravo\n");
        var activity = new DeleteAt
        {
            ContinueOnError = new InArgument<bool>(false),
            Delay = new InArgument<int>(0),
            DeleteLineIndex = new InArgument<int>(1),
        };

        var executed = file.Run<bool>(activity, (a, v) =>
        {
            ((DeleteAt)a).ExecutionResult = new OutArgument<bool>(v);
            return a;
        });

        Assert.True(executed);
        Assert.Equal("bravo\n", file.Contents);
    }

    private static Exception Unwrap(Exception error) =>
        error.InnerException is null ? error : Unwrap(error.InnerException);
}
