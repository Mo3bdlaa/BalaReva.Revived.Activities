# BalaReva Revived — Easy Excel Activities

Excel activities for UiPath, reimplemented on .NET 8.

Type and property names match the published `BalaReva.EasyExcel.Activities` 32.0.0, so
existing workflows keep binding after the swap. A test suite checks the surface back
against the original on every build, including all 27 enums — `FileFormatEnum` and
`SaveEnum` carry 54 formats each, and `PaperSizeEnum` 43 paper sizes.

This package targets `net8.0-windows` and, for all but three of its 77 activities,
requires Excel to be installed.

## What runs without Excel

`GetHiddenRows` and `GetHiddenColumns` read the workbook's Open XML package directly
through `DocumentFormat.OpenXml`, so they need no Excel and no scope. The published
package did the same, which is why they sit on their own base and carry their own file
path. `GetColumnName` and `GetColumnNumber` are arithmetic on Excel's bijective base-26
column letters and need nothing at all.

Everything else goes through COM, behind `IExcelService`, `IExcelWorkbook` and
`IOpenXmlReader`, so a workflow can register its own implementation as an extension.

## The one interop type in the binding surface

`SetBorder.LineStyle` is a `Microsoft.Office.Interop.Excel.XlLineStyle`, exactly as the
published package declared it. A workflow binds by type, so it stays.

`ExcelParam.ExcelWorkBook` is the one place this package does not match. It is a
`Workbook` there and an `object` here, so a workflow reaching for the live COM workbook
has to cast. The reason is that every Office interop assembly on NuGet has a hard
reference on `office` (`Microsoft.Office.Core`), which Microsoft publishes nowhere: the
CLR goes looking for it the moment it loads a member typed that way, and since the scope
hands this object to its body, typing it `Workbook` would stop the whole scope loading on
a machine without `office.dll`. The deviation is recorded, with its reason, in
`audit/verify_binding_surface.py`.

The same constraint is why a few members inside `ExcelWorkbook` go through `dynamic`:
`AutomationSecurity`, `CommandBars.AdaptiveMenus`, `Shape.Type` and `LockAspectRatio` are
all typed in terms of `Microsoft.Office.Core`. They take the numbers directly, and the
numbers are named in `Mso`.

`ChartEmbedToPowerPoint` drives PowerPoint late-bound through its ProgID, for the same
reason and as `BalaReva.Revived.EasyPowerPoint.Activities` does throughout.

## What is not covered by tests

**No CI agent has Excel installed**, so `ExcelService` and `ExcelWorkbook` — the whole COM
half — are not exercised by any automated test. The activity layer above them is, against
a recording stand-in.

`OpenXmlReader` is the exception, and it is tested properly: its tests write real `.xlsx`
files with hidden rows and column spans and read them back. They still only run on the
Windows CI job, because the package targets `net8.0-windows` and so does its test
assembly, which needs a Windows Desktop runtime to host. The reader's own code is
portable; its tests are not.

## Oddities carried over because workflows bind to them

- `TrueFaleNoneEnum` is misspelled in the published package. It stays.
- `InsertHyperlink.DisplayTextOverwirte` is misspelled. It stays.
- `Settings.General.Adaptive_Menus` and `TabColor.Tab_Color` keep their underscores.
- `CountARange` and `CountRange` are different activities: one counts non-empty cells,
  the other counts numbers.
- `Charts.ChartDeleteAll` and `Sheets.DeleteAllCharts` do the same thing under two names,
  in two namespaces. Both are published, so both are here.
