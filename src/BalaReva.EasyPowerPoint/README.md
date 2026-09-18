# BalaReva Revived — Easy PowerPoint Activities

PowerPoint activities for UiPath, reimplemented on .NET 8.

Type and property names match the published `BalaReva.EasyPowerPoint.Activities` 11.0.0,
so existing workflows keep binding after the swap. A test suite checks the surface back
against the original on every build, including all 14 enums — `EntryEffectEnum` alone
has 189 members.

This package targets `net8.0-windows` and requires PowerPoint to be installed.

**It carries no Office interop dependency at all.** PowerPoint's object model is
saturated with `Microsoft.Office.Core` types, which Microsoft does not publish on NuGet;
every package that does is a third-party repackage, and depending on one would be the
supply-chain problem [docs/AUDIT.md](https://github.com/Mo3bdlaa/BalaReva.Revived.Activities/blob/main/docs/AUDIT.md)
calls out. PowerPoint is driven through its ProgID and late binding instead.

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
