using System.Data;
using System.Printing;
using BalaReva.Printer.Enums;

namespace BalaReva.Printer;

/// <summary>
/// <see cref="IPrinterService"/> on top of WPF's printing stack.
/// </summary>
/// <remarks>
/// Windows only, and unavoidably so: this is the same API the published package used,
/// and the package's own public types (PrinterStatus, AccessRightsEnum) are shaped by
/// it. Every method here is a thin translation, which is what keeps the logic that is
/// worth testing in the activities instead.
/// </remarks>
public sealed class WindowsPrinterService : IPrinterService
{
    /// <summary>The shared instance used when a workflow registers no extension.</summary>
    public static WindowsPrinterService Instance { get; } = new();

    /// <inheritdoc />
    public string GetDefaultPrinterName()
    {
        using var server = new LocalPrintServer();
        using var queue = server.DefaultPrintQueue;
        return queue.FullName;
    }

    /// <inheritdoc />
    public PrinterStatus GetDefaultPrinterStatus()
    {
        using var server = new LocalPrintServer();
        using var queue = server.DefaultPrintQueue;
        queue.Refresh();
        return Snapshot(queue);
    }

    /// <inheritdoc />
    public string[] GetLocalPrinters() => Names(EnumeratedPrintQueueTypes.Local);

    /// <inheritdoc />
    public string[] GetNetworkPrinters() => Names(EnumeratedPrintQueueTypes.Connections);

    /// <inheritdoc />
    public void Pause(string printerName, AccessRightsEnum desiredAccess)
    {
        using var queue = Open(printerName, desiredAccess, AccessRightsEnum.AdministratePrinter);
        queue.Pause();
        queue.Commit();
    }

    /// <inheritdoc />
    public void Resume(string printerName, AccessRightsEnum desiredAccess)
    {
        using var queue = Open(printerName, desiredAccess, AccessRightsEnum.AdministratePrinter);
        queue.Resume();
        queue.Commit();
    }

    /// <inheritdoc />
    public void ClearQueue(string printerName, AccessRightsEnum desiredAccess)
    {
        using var queue = Open(printerName, desiredAccess, AccessRightsEnum.AdministratePrinter);
        queue.Purge();
        queue.Commit();
    }

    /// <inheritdoc />
    public void SetDefaultPrinter(string printerName)
    {
        using var server = new LocalPrintServer();
        using var queue = server.GetPrintQueue(printerName);
        server.DefaultPrintQueue = queue;
        server.Commit();
    }

    /// <inheritdoc />
    public DataTable GetQueueData(string printerName)
    {
        using var queue = Open(printerName, AccessRightsEnum.None, AccessRightsEnum.UsePrinter);
        queue.Refresh();

        var table = QueueTable();
        foreach (var job in queue.GetPrintJobInfoCollection())
        {
            using (job)
            {
                table.Rows.Add(
                    job.JobIdentifier, job.Name, job.Submitter, job.JobStatus.ToString(),
                    job.NumberOfPages, job.JobSize, job.TimeJobSubmitted, job.PositionInPrintQueue);
            }
        }
        return table;
    }

    /// <inheritdoc />
    public void RunCommand(CommandEnum command)
    {
        using var server = new LocalPrintServer();
        using var queue = server.DefaultPrintQueue;
        switch (command)
        {
            case CommandEnum.Pause: queue.Pause(); break;
            case CommandEnum.Resume: queue.Resume(); break;
            case CommandEnum.RemoveAllJobs: queue.Purge(); break;
            case CommandEnum.Refresh: queue.Refresh(); return; // Refresh reads; nothing to commit.
            default:
                throw new ArgumentOutOfRangeException(nameof(command), command, "Unknown printer command.");
        }
        queue.Commit();
    }

    /// <summary>The shape of the table <see cref="GetQueueData"/> returns.</summary>
    /// <remarks>Public so a stand-in implementation can produce the same columns.</remarks>
    public static DataTable QueueTable()
    {
        var table = new DataTable("PrintQueue");
        table.Columns.Add("JobId", typeof(int));
        table.Columns.Add("JobName", typeof(string));
        table.Columns.Add("Submitter", typeof(string));
        table.Columns.Add("Status", typeof(string));
        table.Columns.Add("Pages", typeof(int));
        table.Columns.Add("SizeInBytes", typeof(int));
        table.Columns.Add("SubmittedAt", typeof(DateTime));
        table.Columns.Add("PositionInQueue", typeof(int));
        return table;
    }

    private static string[] Names(EnumeratedPrintQueueTypes type)
    {
        using var server = new LocalPrintServer();
        var queues = server.GetPrintQueues([type]);
        var names = new List<string>();
        foreach (var queue in queues)
        {
            using (queue) names.Add(queue.FullName);
        }
        return [.. names];
    }

    /// <summary>
    /// Opens a queue, falling back to <paramref name="fallback"/> when the workflow
    /// left <c>DesiredAccess</c> at <see cref="AccessRightsEnum.None"/>.
    /// </summary>
    private static PrintQueue Open(string printerName, AccessRightsEnum desired, AccessRightsEnum fallback)
    {
        var access = desired == AccessRightsEnum.None ? fallback : desired;
        return new PrintQueue(new LocalPrintServer(), printerName, (PrintSystemDesiredAccess)access);
    }

    private static PrinterStatus Snapshot(PrintQueue q) => new()
    {
        FullName = q.FullName,
        HasPaperProblem = q.HasPaperProblem,
        HasToner = q.HasToner,
        InPartialTrust = q.InPartialTrust,
        IsBidiEnabled = q.IsBidiEnabled,
        IsBusy = q.IsBusy,
        IsDevQueryEnabled = q.IsDevQueryEnabled,
        IsDirect = q.IsDirect,
        IsDoorOpened = q.IsDoorOpened,
        IsHidden = q.IsHidden,
        IsInError = q.IsInError,
        IsInitializing = q.IsInitializing,
        IsIOActive = q.IsIOActive,
        IsManualFeedRequired = q.IsManualFeedRequired,
        IsNotAvailable = q.IsNotAvailable,
        IsOffline = q.IsOffline,
        IsOutOfMemory = q.IsOutOfMemory,
        IsOutOfPaper = q.IsOutOfPaper,
        IsOutputBinFull = q.IsOutputBinFull,
        IsPaperJammed = q.IsPaperJammed,
        IsPaused = q.IsPaused,
        IsPendingDeletion = q.IsPendingDeletion,
        IsPowerSaveOn = q.IsPowerSaveOn,
        IsPrinting = q.IsPrinting,
        IsProcessing = q.IsProcessing,
        IsPublished = q.IsPublished,
        IsQueued = q.IsQueued,
        IsRawOnlyEnabled = q.IsRawOnlyEnabled,
        IsServerUnknown = q.IsServerUnknown,
        IsShared = q.IsShared,
        IsTonerLow = q.IsTonerLow,
        IsWaiting = q.IsWaiting,
        IsWarmingUp = q.IsWarmingUp,
        IsXpsDevice = q.IsXpsDevice,
        KeepPrintedJobs = q.KeepPrintedJobs,
        NeedUserIntervention = q.NeedUserIntervention,
        PagePunt = q.PagePunt,
        PrintingIsCancelled = q.PrintingIsCancelled,
        ScheduleCompletedJobsFirst = q.ScheduleCompletedJobsFirst,
    };
}
