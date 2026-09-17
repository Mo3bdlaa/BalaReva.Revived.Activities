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
    public void With_no_extension_registered_the_activity_reaches_for_the_real_spooler()
    {
        // No IPrinterService extension, so BaseActivity falls back to the Windows
        // implementation. On a machine with no print subsystem that throws, which is
        // the point: the fallback is wired up rather than silently doing nothing.
        var activity = new GetDefaultPrinter { ContinueOnError = new InArgument<bool>(true) };

        var outputs = new WorkflowInvoker(activity).Invoke();

        Assert.False((bool)outputs["ExecutionResult"]);
    }
}
