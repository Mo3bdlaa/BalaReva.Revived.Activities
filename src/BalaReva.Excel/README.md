# BalaReva Revived — Excel Activities

Excel activities for UiPath, reimplemented on .NET 8.

Type and property names match the published `BalaReva.Excel.Activities` 2021.1.0, so
existing workflows keep binding after the swap. A test suite checks the surface back
against the original on every build, including all 23 enums.

This package targets `net8.0-windows` and requires Excel to be installed. Its binding
surface leaks no Excel types, so `DocumentFormat.OpenXml` was on the table, but the
activity list is not: charts built from `XlChartType`, `InsertTableFormat`'s 61
`XlRangeAutoFormat` styles, `ExportWorkBook` to PDF, `ClipboardToDatatable` and the
AutoFit activities all need Excel running.

Unlike the package it replaces, the Excel interop assembly is a **declared NuGet
dependency** rather than a loose DLL copied into `lib/`.

Excel access sits behind `IExcelService` and `IExcelWorkbook`, so a workflow can
register its own implementation as an extension. There is no scope activity: each
activity carries its own file path, sheet name and passwords, opens the workbook, does
its work and closes it again — which is how the published package was shaped.

**No CI agent has Excel installed**, so the COM layer of this package is not covered by
automated tests — only the activity layer above it is.

Two misspellings are carried over from the published package because workflows bind to
them: the comment base class is `BaseCommnet`, and `BalaReva.Excel.Enums.TextOrientationEumn`
sits alongside a correctly spelled `BalaReva.Excel.Utilities.TextOrientationEnum`.
