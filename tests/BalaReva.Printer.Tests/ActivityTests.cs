using System.Activities;
using System.Data;
using BalaReva.Printer.Enums;

namespace BalaReva.Printer.Tests;

/// <summary>
/// Each activity runs as a real workflow against a stand-in spooler, so argument
/// handling and output mapping are covered without touching a real print queue.
/// </summary>
public class ActivityTests
{
    [Fact]
    public void GetDefaultPrinter_returns_the_spooler_default()
    {
        var service = new FakePrinterService { DefaultPrinterName = "HP LaserJet" };

        var outputs = Harness.Run(new GetDefaultPrinter(), service);

        Assert.Equal("HP LaserJet", outputs["Output"]);
        Assert.Equal(["GetDefaultPrinterName"], service.Calls);
    }

    [Fact]
    public void DefaultPrinterStatus_passes_the_snapshot_straight_through()
    {
        var status = new PrinterStatus { FullName = "HP LaserJet", IsTonerLow = true, IsPaused = true };
        var service = new FakePrinterService { Status = status };

        var outputs = Harness.Run(new DefaultPrinterStatus(), service);

        var result = Assert.IsType<PrinterStatus>(outputs["PrinterResult"]);
        Assert.Equal("HP LaserJet", result.FullName);
        Assert.True(result.IsTonerLow);
        Assert.True(result.IsPaused);
    }

    [Fact]
    public void LocalPrinters_and_NetworkPrinters_read_different_sources()
    {
        var service = new FakePrinterService();

        var local = Harness.Run(new LocalPrinters(), service);
        var network = Harness.Run(new NetworkPrinters(), service);

        Assert.Equal(["Local A", "Local B"], (string[])local["Printers"]);
        Assert.Equal(["\\\\server\\Shared"], (string[])network["Printers"]);
        Assert.Equal(["GetLocalPrinters", "GetNetworkPrinters"], service.Calls);
    }

    [Fact]
    public void SetAsDefaultPrinter_passes_the_name_through()
    {
        var service = new FakePrinterService();

        Harness.Run(new SetAsDefaultPrinter { PrinterName = new InArgument<string>("HP") }, service);

        Assert.Equal(["SetDefaultPrinter(HP)"], service.Calls);
    }

    [Theory]
    [InlineData(AccessRightsEnum.None)]
    [InlineData(AccessRightsEnum.AdministratePrinter)]
    public void PausePrinter_forwards_the_requested_access(AccessRightsEnum access)
    {
        var service = new FakePrinterService();

        Harness.Run(
            new PausePrinter { PrinterName = new InArgument<string>("HP"), DesiredAccess = access },
            service);

        Assert.Equal([$"Pause(HP,{access})"], service.Calls);
    }

    [Fact]
    public void ResumePrinter_forwards_the_requested_access()
    {
        var service = new FakePrinterService();

        Harness.Run(
            new ResumePrinter
            {
                PrinterName = new InArgument<string>("HP"),
                DesiredAccess = AccessRightsEnum.AdministratePrinter,
            },
            service);

        Assert.Equal(["Resume(HP,AdministratePrinter)"], service.Calls);
    }

    [Fact]
    public void ClearPrinterQueue_forwards_the_requested_access()
    {
        var service = new FakePrinterService();

        Harness.Run(new ClearPrinterQueue { PrinterName = new InArgument<string>("HP") }, service);

        Assert.Equal(["ClearQueue(HP,None)"], service.Calls);
    }

    [Theory]
    [InlineData(CommandEnum.Pause)]
    [InlineData(CommandEnum.Resume)]
    [InlineData(CommandEnum.Refresh)]
    [InlineData(CommandEnum.RemoveAllJobs)]
    public void PrinterCommand_forwards_the_command(CommandEnum command)
    {
        var service = new FakePrinterService();

        Harness.Run(new PrinterCommand { Command = command }, service);

        Assert.Equal([$"RunCommand({command})"], service.Calls);
    }

    [Fact]
    public void QueueData_reports_the_table_and_counts_its_rows()
    {
        var table = WindowsPrinterService.QueueTable();
        table.Rows.Add(1, "Report.pdf", "alice", "Spooling", 3, 2048, DateTime.UtcNow, 1);
        table.Rows.Add(2, "Invoice.docx", "bob", "Printing", 1, 512, DateTime.UtcNow, 2);
        var service = new FakePrinterService { Queue = table };

        var outputs = Harness.Run(
            new QueueData { PrinterName = new InArgument<string>("HP") }, service);

        Assert.Same(table, outputs["QueueTable"]);
        Assert.Equal(2, outputs["TotalJobs"]);
    }

    [Fact]
    public void QueueData_reports_zero_for_an_idle_printer()
    {
        var service = new FakePrinterService { Queue = WindowsPrinterService.QueueTable() };

        var outputs = Harness.Run(
            new QueueData { PrinterName = new InArgument<string>("HP") }, service);

        Assert.Equal(0, outputs["TotalJobs"]);
        Assert.Empty(((DataTable)outputs["QueueTable"]).Rows);
    }
}
