namespace BalaReva.Printer;

/// <summary>Operation for <see cref="PrinterCommand"/> to run on the default printer.</summary>
/// <remarks>Names and values match the published package and must not be renumbered.</remarks>
public enum CommandEnum
{
    /// <summary>Pause the queue.</summary>
    Pause = 1,
    /// <summary>Cancel every job in the queue.</summary>
    RemoveAllJobs = 2,
    /// <summary>Refresh the queue's state from the spooler.</summary>
    Refresh = 3,
    /// <summary>Resume a paused queue.</summary>
    Resume = 4,
}
