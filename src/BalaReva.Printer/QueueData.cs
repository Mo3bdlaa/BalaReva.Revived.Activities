using System.Activities;
using System.ComponentModel;
using System.Data;

namespace BalaReva.Printer;

/// <summary>Reads the jobs queued on a printer into a DataTable.</summary>
[DisplayName("Queue Data")]
[Description("Reads the jobs waiting on a printer, one row per job.")]
public sealed class QueueData : BaseActivity
{
    /// <summary>Printer to read the queue of.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Printer Name")]
    [Description("Name of the printer whose queue to read.")]
    public InArgument<string> PrinterName { get; set; } = null!;

    /// <summary>One row per queued job.</summary>
    [Category("Output")]
    [DisplayName("Queue Table")]
    [Description("One row per queued job: JobId, JobName, Submitter, Status, Pages, "
                 + "SizeInBytes, SubmittedAt and PositionInQueue.")]
    public OutArgument<DataTable> QueueTable { get; set; } = null!;

    /// <summary>Number of jobs on the queue.</summary>
    [Category("Output")]
    [DisplayName("Total Jobs")]
    [Description("Number of jobs waiting on the queue.")]
    public OutArgument<int> TotalJobs { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPrinterService service)
    {
        var table = service.GetQueueData(RequirePrinterName(context, PrinterName));
        QueueTable.Set(context, table);
        TotalJobs.Set(context, table.Rows.Count);
    }
}
