# BalaReva Revived — Printer Activities

Printer and print queue activities for UiPath, reimplemented on .NET 8.

Type and property names match the published `BalaReva.Printer.Activities` 2019.2.1, so
existing workflows keep binding after the swap. A test suite checks the surface back
against the original on every build.

This package targets `net8.0-windows` and that is forced by the binding surface, not
chosen: `PrinterStatus` exposes 38 booleans that are `System.Printing.PrintQueue`
members one for one, and `AccessRightsEnum`'s values are `PrintSystemDesiredAccess`
exactly. That is WPF's printing stack, which exists only on Windows.

The spooler sits behind `IPrinterService`, so argument handling, `ContinueOnError` and
output mapping are covered by tests on any machine, and a workflow can register its own
implementation as an extension. With none registered, the activities use the real one.

Activities: `GetDefaultPrinter`, `DefaultPrinterStatus`, `LocalPrinters`,
`NetworkPrinters`, `SetAsDefaultPrinter`, `PausePrinter`, `ResumePrinter`,
`ClearPrinterQueue`, `PrinterCommand`, `QueueData`.

See [docs/REVIVAL.md](https://github.com/Mo3bdlaa/BalaReva.Revived.Activities/blob/main/docs/REVIVAL.md).
