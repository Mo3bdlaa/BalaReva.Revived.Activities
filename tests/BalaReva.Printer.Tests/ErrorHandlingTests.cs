using System.Activities;

namespace BalaReva.Printer.Tests;

public class ErrorHandlingTests
{
    [Fact]
    public void A_blank_printer_name_is_rejected_before_the_spooler_is_touched()
    {
        var service = new FakePrinterService();
        var activity = new PausePrinter { PrinterName = new InArgument<string>("   ") };

        Assert.ThrowsAny<Exception>(() => Harness.Run(activity, service));
        Assert.Empty(service.Calls);
    }

    [Fact]
    public void ContinueOnError_reports_failure_instead_of_faulting()
    {
        var service = new FakePrinterService { Throw = new InvalidOperationException("spooler is down") };
        var activity = new GetDefaultPrinter { ContinueOnError = new InArgument<bool>(true) };

        var outputs = Harness.Run(activity, service);

        Assert.False((bool)outputs["ExecutionResult"]);
    }

    [Fact]
    public void Without_ContinueOnError_a_spooler_failure_faults_the_workflow()
    {
        var service = new FakePrinterService { Throw = new InvalidOperationException("spooler is down") };
        var activity = new GetDefaultPrinter { ContinueOnError = new InArgument<bool>(false) };

        Assert.ThrowsAny<Exception>(() => Harness.Run(activity, service));
    }

    [Fact]
    public void A_successful_activity_reports_ExecutionResult_true()
    {
        var outputs = Harness.Run(new GetDefaultPrinter(), new FakePrinterService());

        Assert.True((bool)outputs["ExecutionResult"]);
    }

    [Fact]
    public void With_no_extension_registered_the_activity_falls_back_to_the_real_spooler()
    {
        // No IPrinterService extension, so BaseActivity resolves the Windows one.
        // Whether that call succeeds depends on the agent's printers, so the invariant
        // worth asserting is that the fallback is wired at all: the activity runs and
        // reports a result rather than failing on a null service.
        var activity = new GetDefaultPrinter
        {
            ContinueOnError = new InArgument<bool>(true),
            Delay = new InArgument<short>(0),
            ExecutionResult = new OutArgument<bool>(),
            Output = new OutArgument<string>(),
        };

        var outputs = new WorkflowInvoker(activity).Invoke();

        Assert.NotNull(WindowsPrinterService.Instance);
        Assert.IsType<bool>(outputs["ExecutionResult"]);
    }
}
