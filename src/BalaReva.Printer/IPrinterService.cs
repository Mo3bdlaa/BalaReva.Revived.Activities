using System.Data;
using BalaReva.Printer.Enums;

namespace BalaReva.Printer;

/// <summary>
/// The print spooler operations the activities need.
/// </summary>
/// <remarks>
/// The activities talk to this rather than to <c>System.Printing</c> directly, so that
/// argument handling, ContinueOnError and output mapping can be tested on any machine
/// with a stand-in. The Windows implementation behind it stays thin enough to be worth
/// reviewing by eye.
///
/// A workflow can supply its own through <c>WorkflowInvoker.Extensions</c>; with none
/// registered the activities fall back to <see cref="WindowsPrinterService"/>.
/// </remarks>
public interface IPrinterService
{
    /// <summary>Name of the machine's default printer.</summary>
    string GetDefaultPrinterName();

    /// <summary>A snapshot of the default printer's state.</summary>
    PrinterStatus GetDefaultPrinterStatus();

    /// <summary>Names of the queues installed locally.</summary>
    string[] GetLocalPrinters();

    /// <summary>Names of the queues that come from a print server connection.</summary>
    string[] GetNetworkPrinters();

    /// <summary>Pauses a queue.</summary>
    void Pause(string printerName, AccessRightsEnum desiredAccess);

    /// <summary>Resumes a paused queue.</summary>
    void Resume(string printerName, AccessRightsEnum desiredAccess);

    /// <summary>Cancels every job on a queue.</summary>
    void ClearQueue(string printerName, AccessRightsEnum desiredAccess);

    /// <summary>Makes a printer the machine default.</summary>
    void SetDefaultPrinter(string printerName);

    /// <summary>The jobs queued on a printer, one row each.</summary>
    DataTable GetQueueData(string printerName);

    /// <summary>Runs a queue operation against the default printer.</summary>
    void RunCommand(CommandEnum command);
}
