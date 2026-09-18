# BalaReva Revived — Word Activities

Word activities for UiPath, reimplemented on .NET 8.

Type and property names match the published `BalaReva.Word.Activities` 8.0.0, so
existing workflows keep binding after the swap. A test suite checks the surface back
against the original on every build, including all 19 enums.

This package targets `net8.0-windows` and requires Word to be installed. Unlike Excel
and PowerPoint, Word could not be moved to `DocumentFormat.OpenXml`: `ExecuteMacro`,
`Paste`, `CopyTableToClipboard`, `PrintDocument`, `WordToPdf`, `SaveAs` across 25 Word
formats, `WordStatistics`' page count and `CloseAllWord` all need Word itself running.

Unlike the package it replaces, the Word interop assembly is a **declared NuGet
dependency** rather than a loose DLL copied into `lib/`, where no manifest-based scanner
can see it.

Word access sits behind `IWordService` and `IWordDocument`, so a workflow can register
its own implementation as an extension.

**No CI agent has Word installed**, so the COM layer of this package is not covered by
automated tests — only the activity layer above it is. See
[docs/REVIVAL.md](https://github.com/Mo3bdlaa/BalaReva.Revived.Activities/blob/main/docs/REVIVAL.md).

Two oddities carried over from the published package, because workflows bind to them:
`WordObject.ModiPassword` is spelled that way, and `Documents.FindAndReplace` takes only
two password arguments — it is a password change that was misnamed at the source.
