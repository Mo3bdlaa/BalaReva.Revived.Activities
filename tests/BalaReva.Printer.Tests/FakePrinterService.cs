using System.Data;
using BalaReva.Printer.Enums;

namespace BalaReva.Printer.Tests;

/// <summary>
/// A stand-in spooler that records what it was asked to do.
/// </summary>
/// <remarks>
/// Real queue operations cannot be tested: pausing a queue or purging jobs on a build
/// agent is destructive, and an agent may have no printers at all. Everything worth
/// testing in these activities is on this side of the boundary anyway — argument
/// handling, which service call is made with which arguments, and how results are
/// mapped back onto output arguments.
/// </remarks>
public sealed class FakePrinterService : IPrinterService
{
    public List<string> Calls { get; } = [];

    public string DefaultPrinterName { get; set; } = "Fake Printer";

    public PrinterStatus Status { get; set; } = new() { FullName = "Fake Printer", IsPaused = true };

    public string[] Local { get; set; } = ["Local A", "Local B"];

    public string[] Network { get; set; } = ["\\\\server\\Shared"];

    public DataTable Queue { get; set; } = WindowsPrinterService.QueueTable();

    public Exception? Throw { get; set; }

    private T Record<T>(string call, T result)
    {
        Calls.Add(call);
        if (Throw is not null) throw Throw;
        return result;
    }

    public string GetDefaultPrinterName() => Record("GetDefaultPrinterName", DefaultPrinterName);

    public PrinterStatus GetDefaultPrinterStatus() => Record("GetDefaultPrinterStatus", Status);

    public string[] GetLocalPrinters() => Record("GetLocalPrinters", Local);

    public string[] GetNetworkPrinters() => Record("GetNetworkPrinters", Network);

    public void Pause(string printerName, AccessRightsEnum access) =>
        Record($"Pause({printerName},{access})", 0);

    public void Resume(string printerName, AccessRightsEnum access) =>
        Record($"Resume({printerName},{access})", 0);

    public void ClearQueue(string printerName, AccessRightsEnum access) =>
        Record($"ClearQueue({printerName},{access})", 0);

    public void SetDefaultPrinter(string printerName) =>
        Record($"SetDefaultPrinter({printerName})", 0);

    public DataTable GetQueueData(string printerName) =>
        Record($"GetQueueData({printerName})", Queue);

    public void RunCommand(CommandEnum command) => Record($"RunCommand({command})", 0);
}
