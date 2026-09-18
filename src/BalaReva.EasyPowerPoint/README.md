# BalaReva Revived — Easy PowerPoint Activities

PowerPoint activities for UiPath, reimplemented on .NET 8.

Type and property names match the published `BalaReva.EasyPowerPoint.Activities` 11.0.0,
so existing workflows keep binding after the swap. A test suite checks the surface back
against the original on every build, including all 14 enums — `EntryEffectEnum` alone
has 189 members.

This package targets `net8.0-windows` and requires PowerPoint to be installed.

**PowerPoint is driven late-bound, through its ProgID.** Its object model is saturated
with `Microsoft.Office.Core` types — `MsoTriState` alone turns up on picture insertion,
text orientation, shape ordering and the fixed-format export — and Microsoft publishes
no `Office.dll` on NuGet. The one package that carries it is an Office 2007-era
third-party repackage with the assembly sitting at the archive root, which is exactly
the provenance [docs/AUDIT.md](https://github.com/Mo3bdlaa/BalaReva.Revived.Activities/blob/main/docs/AUDIT.md)
exists to flag. Late binding needs none of it.

**It carries no Office interop dependency at all**, and referencing one is not the easy
fix it looks like. Every Office interop assembly on NuGet — PowerPoint's included — has a
hard reference on `office` (`Microsoft.Office.Core`), which Microsoft publishes nowhere.
The reference compiles; the CLR then goes looking for `office.dll` the moment it loads a
member typed that way, and does not find it. Trying this cost a CI run: 26 tests failed
with `Could not load file or assembly 'office'`, and not only on the property concerned —
the scope hands that object to its body, so the scope itself stopped working.

So one thing here does not match the published package: `PowerPointObject.PptPersentation`
is a `Presentation` there and an `object` here. A workflow reaching for the live COM
presentation has to cast. That is the whole of the difference, it is recorded as an
accepted deviation in `audit/verify_binding_surface.py`, and it buys a scope that runs on
a machine that has PowerPoint but not the Office PIAs — which is most of them.

PowerPoint access sits behind `IPowerPointService` and `IPowerPointPresentation`, so a
workflow can register its own implementation as an extension.

**No CI agent has PowerPoint installed**, so the COM layer is not covered by automated
tests — only the activity layer above it is, and the late-bound service has no
compile-time checking either. Treat `PowerPointService` as the least-verified code here.

Oddities carried over because workflows bind to them:

- `ExtractHyperLinks` and `GetRowItem` live under `BalaReva.Easy.PowerPoint` — note the
  extra dot — while the other 54 activities are under `BalaReva.EasyPowerPoint`.
- `ImageShapeCount` reports through an output called `ChartCount`, and `TextShapeCount`
  through one called `ImageCount`. They are named the other way round at the source.
- `Find_Replace` keeps its underscore.

`ExportTableToExcel` and `ImportDataFromExcel` are not implemented: they need the Excel
object model, which this package deliberately does not depend on. Both throw with a
message pointing at `BalaReva.Revived.Excel.Activities`.
