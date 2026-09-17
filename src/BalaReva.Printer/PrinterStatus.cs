namespace BalaReva.Printer;

/// <summary>
/// A snapshot of a print queue's state.
/// </summary>
/// <remarks>
/// Every member here mirrors a <c>System.Printing.PrintQueue</c> property of the same
/// name, which is how the published package exposed it. It is a snapshot rather than a
/// live view: reading a property does not go back to the spooler.
/// </remarks>
public sealed class PrinterStatus
{
    /// <summary>Full name of the queue, as the print server reports it.</summary>
    public string FullName { get; set; } = string.Empty;

    /// <summary>The printer is reporting a paper problem.</summary>
    public bool HasPaperProblem { get; set; }

    /// <summary>The printer reports that it has toner.</summary>
    public bool HasToner { get; set; }

    /// <summary>The queue is being accessed in partial trust.</summary>
    public bool InPartialTrust { get; set; }

    /// <summary>Two-way communication with the printer is enabled.</summary>
    public bool IsBidiEnabled { get; set; }

    /// <summary>The printer is busy.</summary>
    public bool IsBusy { get; set; }

    /// <summary>The queue holds jobs whose settings do not match the printer.</summary>
    public bool IsDevQueryEnabled { get; set; }

    /// <summary>Jobs print directly rather than being spooled.</summary>
    public bool IsDirect { get; set; }

    /// <summary>A door or cover on the printer is open.</summary>
    public bool IsDoorOpened { get; set; }

    /// <summary>The queue is hidden from users.</summary>
    public bool IsHidden { get; set; }

    /// <summary>The printer is in an error state.</summary>
    public bool IsInError { get; set; }

    /// <summary>The printer is initializing.</summary>
    public bool IsInitializing { get; set; }

    /// <summary>The printer is exchanging data with the server.</summary>
    public bool IsIOActive { get; set; }

    /// <summary>The printer is waiting for paper to be fed by hand.</summary>
    public bool IsManualFeedRequired { get; set; }

    /// <summary>The printer is not available.</summary>
    public bool IsNotAvailable { get; set; }

    /// <summary>The printer is offline.</summary>
    public bool IsOffline { get; set; }

    /// <summary>The printer is out of memory.</summary>
    public bool IsOutOfMemory { get; set; }

    /// <summary>The printer is out of paper.</summary>
    public bool IsOutOfPaper { get; set; }

    /// <summary>The output bin is full.</summary>
    public bool IsOutputBinFull { get; set; }

    /// <summary>Paper is jammed in the printer.</summary>
    public bool IsPaperJammed { get; set; }

    /// <summary>The queue is paused.</summary>
    public bool IsPaused { get; set; }

    /// <summary>The queue is being deleted.</summary>
    public bool IsPendingDeletion { get; set; }

    /// <summary>The printer is in power save mode.</summary>
    public bool IsPowerSaveOn { get; set; }

    /// <summary>The printer is printing.</summary>
    public bool IsPrinting { get; set; }

    /// <summary>The queue is processing a job.</summary>
    public bool IsProcessing { get; set; }

    /// <summary>The queue is published to the directory service.</summary>
    public bool IsPublished { get; set; }

    /// <summary>The printer accepts more than one job at a time.</summary>
    public bool IsQueued { get; set; }

    /// <summary>The queue accepts only raw data jobs.</summary>
    public bool IsRawOnlyEnabled { get; set; }

    /// <summary>The printer is in an error state of unknown cause.</summary>
    public bool IsServerUnknown { get; set; }

    /// <summary>The queue is shared.</summary>
    public bool IsShared { get; set; }

    /// <summary>The printer is low on toner.</summary>
    public bool IsTonerLow { get; set; }

    /// <summary>The queue is waiting for a job.</summary>
    public bool IsWaiting { get; set; }

    /// <summary>The printer is warming up.</summary>
    public bool IsWarmingUp { get; set; }

    /// <summary>The printer understands XPS directly.</summary>
    public bool IsXpsDevice { get; set; }

    /// <summary>Printed jobs are kept in the queue.</summary>
    public bool KeepPrintedJobs { get; set; }

    /// <summary>The printer needs someone to attend to it.</summary>
    public bool NeedUserIntervention { get; set; }

    /// <summary>The printer could not print the current page.</summary>
    public bool PagePunt { get; set; }

    /// <summary>Printing of the current job has been cancelled.</summary>
    public bool PrintingIsCancelled { get; set; }

    /// <summary>Fully spooled jobs are scheduled ahead of others.</summary>
    public bool ScheduleCompletedJobsFirst { get; set; }
}
