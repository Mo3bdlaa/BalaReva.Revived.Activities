# Reviving the packages

## Why this is a reimplementation

The published BalaReva packages are closed-source. There is no repository to fork
and no source to retarget, so "revive" cannot mean bumping a `TargetFramework` and
rebuilding. Each activity has to be written again.

What must be preserved is the **binding surface**. A `.xaml` workflow refers to an
activity by its type's full name and to each input and output by property name. Get
a namespace, a type name, a property name or an enum member wrong and the workflow
breaks the moment someone swaps the package — which is the one failure mode a
"revived" package must not have.

So the process is:

1. `tools/ApiSurface` records the published surface: activity types, their
   properties and signatures, and enum members. It reads metadata tables only —
   names and signatures, never method bodies.
2. The activity is written fresh against .NET 8 and current dependencies.
3. `ApiCompatibilityTests` checks the result back against the recorded surface, so
   a rename fails the build rather than a customer's workflow.

## Scope of the current pass

Targeting the eight packages that already ship .NET 6 assets, since .NET 6 left
support on 2024-11-12. Between them they hold **262 concrete activities**:

| Package | Activities | Status |
|---|---:|---|
| `BalaReva.EasyText.Activities` | 12 | ✅ Reimplemented on .NET 8 |
| `BalaReva.EasyImage.Activities` | 9 | Not started |
| `BalaReva.Printer.Activities` | 10 | Not started |
| `BalaReva.EasyOutlook.Activities` | 20 | Not started |
| `BalaReva.Excel.Activities` | 39 | Not started |
| `BalaReva.Word.Activities` | 39 | Not started |
| `BalaReva.EasyPowerPoint.Activities` | 56 | Not started |
| `BalaReva.EasyExcel.Activities` | 77 | Not started |

EasyText went first because it is the only one of the eight with no Windows
dependency at all, which means its behaviour can be **executed and asserted in CI**
rather than merely compiled. Everything after it leans on Office COM interop,
`System.Drawing` or the Windows print spooler, and will need a Windows runner to be
tested honestly.

### A note on the Office-bound packages

EasyExcel, Excel, Word, EasyPowerPoint and EasyOutlook reach Office through COM
interop. Retargeting them to .NET 8 makes them supported again, but it cannot make
them portable: they will still require Windows with a matching Office installation,
and they will still be unusable from a Cross-platform project. Where an activity can
be served by `DocumentFormat.OpenXml` instead of interop, that is worth doing on its
own merits — no Office install, no COM lifetime problems, and it runs anywhere — but
it is a larger change than a retarget and will not always be behaviour-identical.

## What changed in EasyText, deliberately

- **Targets `net8.0`, not `net8.0-windows`.** Nothing in these activities touches
  Windows; they read and write text files. The published package shipped
  `net461` + `net6.0-windows`. Plain `net8.0` keeps them working in a Studio
  *Windows* project and additionally makes them usable from a *Cross-platform*
  project, which the original could not do.
- **Encoding and line endings are preserved.** Every write restores the file's
  original encoding, its dominant line ending, and whether it ended with a trailing
  newline. An activity that rewrites a CRLF file as LF turns one changed line into a
  whole-file diff.
- **No WPF designers.** The published package shipped custom `ActivityDesigner`
  views; this one relies on Studio's generated property grid. That changes how the
  activity looks, not what a workflow binds to, and it is what keeps the runtime
  assembly free of a Windows dependency.

## Behaviour that the metadata could not settle

Type and property names come from the assemblies and are certain. Semantics do not,
and the following were decided rather than discovered. Each needs confirming against
the published package on a machine with Studio, and each is covered by a test that
will make a correction cheap:

- **Line numbers are 1-based.** `LineAt`, `LineFrom`, `LineTo`, `DeleteLineIndex`
  and the results of `FindLineIndex` all count the first line as 1. The name
  `DeleteLineIndex` hints at 0-based, but every other member says "line", and 1-based
  is what an editor shows. Character offsets — `FindText.StartIndex` and its
  `Result` — stay 0-based to match `String.IndexOf`.
- **A trailing newline terminates the last line** rather than starting an empty one,
  so `"a\nb\n"` is two lines, not three.
- **Find and replace is ordinal and case-sensitive.** `FindReplaceAllText` and
  `FindReplaceAt` have no comparison property, unlike `FindText` and `FindLineIndex`
  which do.
- **`RemoveEmptyLine` also removes whitespace-only lines**, not just zero-length ones.
- **`InsertLineAt` accepts line count + 1** as an append.
- **`ContinueOnError` reports through `ExecutionResult`** and leaves the file
  untouched when an activity fails partway.
