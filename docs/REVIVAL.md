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

| Package | Activities | Target | Status |
|---|---:|---|---|
| `BalaReva.EasyText.Activities` | 12 | `net8.0` | ✅ Reimplemented |
| `BalaReva.EasyImage.Activities` | 9 | `net8.0-windows` | ✅ Reimplemented |
| `BalaReva.Printer.Activities` | 10 | `net8.0-windows` | ✅ Reimplemented |
| `BalaReva.EasyOutlook.Activities` | 20 | `net8.0-windows` | ✅ Reimplemented |
| `BalaReva.Excel.Activities` | 39 | `net8.0-windows` | ✅ Reimplemented |
| `BalaReva.Word.Activities` | 39 | `net8.0-windows` | ✅ Reimplemented |
| `BalaReva.EasyPowerPoint.Activities` | 56 | `net8.0-windows` | ✅ Reimplemented |
| `BalaReva.EasyExcel.Activities` | 77 | — | Not started |

EasyText went first because it is the only one of the eight with no Windows
dependency at all, which means its behaviour can be executed and asserted on any
agent rather than merely compiled.

### Why EasyImage, Printer and EasyOutlook have to stay on Windows

EasyText could be moved to plain `net8.0` because nothing in it needed Windows. That
is not a choice available for these two, and the reason is the binding surface itself:

- **EasyImage** exposes `System.Drawing` types as arguments. `ImageRotate.FlipType` is
  a `RotateFlipType`, and `ImageWatermark` takes a `Font`, a `Color`, a `Point` and an
  `ImageFormat`. A workflow's `.xaml` names those types, so substituting a portable
  imaging library such as ImageSharp or SkiaSharp would break exactly the workflows
  this package exists to keep working. `System.Drawing.Common` is Windows-only from
  .NET 7 onward, so the package is Windows-only too.
- **Printer** exposes `PrinterStatus`, whose 38 booleans are `System.Printing.PrintQueue`
  members one for one, and `AccessRightsEnum`, whose values are `PrintSystemDesiredAccess`
  exactly — `AdministratePrinter` is 983052 in both. That is WPF's printing stack, which
  ships only in the Windows Desktop runtime.
- **EasyOutlook** hands back a live `Microsoft.Office.Interop.Outlook.MailItem` through
  `BalaReva.Outlook.EmailItem`, so the COM type is part of the public API. Its three
  enums are Outlook's own numbers too: `MailFolderEnum` is `OlDefaultFolders` (hence the
  non-consecutive 3, 4, 5, 6, 16, 23), `BusyStatusEnum` is `OlBusyStatus` and
  `ImportanceEnum` is `OlImportance`.

Both therefore target `net8.0-windows`. They are supported again, and usable from a
Studio *Windows* project, but they cannot be used from a *Cross-platform* project — and
no amount of retargeting would change that without breaking compatibility.

### How they are tested

The Windows-only packages cannot run on a Linux agent at all: `System.Drawing.Common`
throws there, and a `net8.0-windows` test host needs a Windows Desktop runtime that
does not exist on Linux. So CI has two jobs. The Linux job builds everything (which
works, via `EnableWindowsTargeting`) and runs the EasyText suite. The Windows job runs
all three suites.

Printer talks to the spooler through `IPrinterService`, and EasyOutlook talks to the
mailbox through `IOutlookService`; both register a stand-in as a workflow extension in
their tests. That is not only for portability: pausing a queue or
purging jobs on a build agent is destructive, and an agent may have no printers at all.
Everything worth testing — argument validation, which call is made with which
arguments, `ContinueOnError`, output mapping — sits on this side of that boundary. The
Windows implementation behind it is a thin translation, and is the part that a real
machine has to vouch for.

**The Office-bound packages are the sharpest cases of this and deserve stating plainly:
no CI agent has Outlook, Word or Excel installed, so `OutlookService`, `WordService` and
`ExcelService` — the entire COM half of those packages — are not covered by any
automated test.** It compiles, and the activity layer above it is
well covered, but its behaviour against a real mailbox is unverified. The same will be
true of the Office-bound packages still to come, and it is why the COM layer is kept to
a mechanical translation with every judgement pushed up into the activities.

### Holding the surface

A `.xaml` workflow binds by three things: the type's full name, the property's name, and
the property's type. Change any one and the workflow breaks on upgrade — silently in the
designer, loudly at run time. Each package's test suite asserts all three against the
recorded surface, but those suites are `net8.0-windows`, so they only run on the Windows
CI job, which is the last place to find out.

`audit/verify_binding_surface.py` does the same comparison from the built `.nupkg`,
reading metadata only, so it runs on Linux and on a developer machine. It walks every
public type outside the design assembly, resolves inherited properties up the base
chain — a property we moved onto a shared base is still the same property to a workflow
— and reports every name and type that moved. CI runs it on the Linux job right after
packing.

Building EasyPowerPoint without it cost a red Windows run and 44 differences at once,
among them four `float` arguments declared as `double`, eleven design-time properties
declared as `InArgument<T>`, and four published classes never written at all. All of
them are the kind of thing that looks right in a diff.

### Why Word and Excel stayed on COM

Checking the four Office-bound packages for COM types in their **binding surfaces**
settled what each is free to use underneath. Word and Excel leak none at all: every
enum is their own, so the implementation is invisible to a workflow. EasyExcel leaks
exactly one on an activity, `SetBorder.LineStyle`, which is an
`Microsoft.Office.Interop.Excel.XlLineStyle`.

EasyPowerPoint and EasyExcel each leak one more, and not on an activity: the object the
scope hands its body carries the live COM document — `PowerPointObject.PptPersentation`
is a `Microsoft.Office.Interop.PowerPoint.Presentation`, `ExcelParam.ExcelWorkBook` a
`Microsoft.Office.Interop.Excel.Workbook`. An early pass over the surface missed both
because it read the activities and not the plain classes beside them, which is why the
check described under *Holding the surface* now walks every public type.

Being free to choose is not the same as being able to, and for both packages done so
far the activity lists settle it the other way.

Word needs Word running for `ExecuteMacro`, `Paste`, `CopyTableToClipboard`,
`PrintDocument`, `WordToPdf`, `SaveAs` across 25 Word formats, `WordStatistics`' page
count and `CloseAllWord`. Excel needs Excel running for charts built from `XlChartType`,
`InsertTableFormat`'s 61 `XlRangeAutoFormat` styles, `ExportWorkBook` to PDF,
`ClipboardToDatatable` and the AutoFit activities.

`DocumentFormat.OpenXml` could serve maybe half of each, but a package that worked for
half its own activities would be worse than one that is honest about its dependency. So
both are COM, behind a service interface, with the activity layer tested through a
stand-in.

One activity is left deliberately unimplemented behind a clear exception rather than
quietly doing nothing: `ImageExtract` and `ExtractHeaderFooterImages` round-tripped
images through the Windows clipboard in the published package, and reproducing that
would mean pulling WinForms into a document package. The exception says so and points
here.

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

## What changed in EasyImage and Printer, deliberately

- **Images are loaded through a stream, not `Image.FromFile`.** That method holds a
  lock on the source file for the lifetime of the image, so an activity could not write
  its result back over its own input. Copying through a stream lifts that restriction,
  and there is a test for it.
- **`ImageCropper` rejects a rectangle that runs off the edge.** GDI+ silently clamps
  it and produces a smaller image than asked for, which in a robot surfaces much later
  as a mysteriously wrong file.
- **Destination folders are created if missing**, so a workflow does not have to
  precede every image activity with a folder check.
- **`ImageWatermark` only disposes a font it created itself.** A `Font` handed in by
  the workflow belongs to the workflow and may be reused on the next iteration.
- **`DesiredAccess` of `None` means "whatever the operation needs"** rather than an
  error: pause, resume and purge open the queue with `AdministratePrinter`, and reading
  the queue uses `UsePrinter`.

### `ImageWatermark`'s Font and ImageFormat cannot be literals

This is inherited from the original's argument types rather than introduced here, but
it will bite anyone wiring the activity up, so it is worth stating plainly.

WF's `Literal<T>` accepts only value types and `String`. `ImageWatermark.Font` and
`ImageWatermark.ImageFormat` are `InArgument<Font>` and `InArgument<ImageFormat>`, both
reference types, so a literal binding is rejected at validation time with:

> `Literal only supports value types and the immutable type System.String.`

They have to be bound as expressions instead — in Studio that means a VB or C#
expression such as `New Font("Arial", 24)` rather than a design-time constant. The
other two `System.Drawing` arguments, `ForeColor` and `TextPosition`, are structs and
bind as literals without trouble. Leaving both unset is also fine: the activity falls
back to 24pt bold Arial, semi-transparent white, and the source image's own format.

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

For EasyImage and Printer:

- **`ImageCombine` lays images out horizontally and `ImageMerge` stacks them
  vertically.** The published package ships both taking the same three arguments, and
  metadata cannot say which axis each used. This at least makes the pair distinct and
  useful; if the original did it the other way round, swapping them is a two-line
  change plus the two tests that pin it.
- **The combined canvas takes the larger of the two sizes on the cross axis**, so
  mismatched images are padded rather than stretched or clipped.
- **`ImageConverter.FileFormat` wins over the destination file extension.** Converting
  to PNG with a `.jpg` destination really does write PNG bytes. Silently ignoring the
  property the user set would be worse than the mismatch.
- **`ImageWatermark` names its output** after the source file with an extension from
  the chosen `ImageFormat`, written into `DestinationFolder`.
- **`PrinterCommand` acts on the default printer**, since it has no printer name
  argument. The named-printer operations are the separate Pause, Resume and Clear
  activities.

For EasyOutlook:

- **`GetFolders` lists the folders directly under the mailbox root**, and
  `CreateFolder`, `DeleteFolder` and `RenameFolder` act at that same level. The
  published names say "folder" and "sub folder" without saying what either is relative
  to, so this is the reading that makes the pair consistent.
- **"Empty" means no items and no sub folders**, and the sweep is one level deep, so a
  folder whose only child is itself empty survives `DeleteEmptyFolders`.
- **The SharePoint activities look for a store named "SharePoint Lists".** That is what
  an English Outlook calls it; on a localised install the name differs and the activity
  will report that no such store exists. Worth revisiting if anyone hits it.
- **`ExecuteMacro` goes through late binding.** Outlook's object model, unlike Excel's
  and Word's, has no documented `Application.Run`, so there is no clean way to do this
  and no way to verify it here. If it fails, the COM error is surfaced rather than
  swallowed.
- **A meeting sends, an appointment saves.** `NewAppointment` writes to the calendar;
  `NewMeeting` sends invitations as soon as it runs.
- **`ReminderMinutes` of zero or less means no reminder** rather than "remind at the
  start", since setting the minutes without also setting `ReminderSet` does nothing.
- **An all-day event ignores the times**, so its end may precede its start on the clock
  without being rejected.

For Word:

- **`Documents.FindAndReplace` is a password change.** Its only two arguments are
  `NewOpenPassword` and `NewModifyPassword`, which is what `ChangePassword` takes; it
  looks like a copy that was renamed and shipped. Renaming it here would break any
  workflow bound to it, so it keeps the name and does what its arguments say. Real find
  and replace is `Pages.FindReplace`, which has the arguments for it.
- **`WordObject.ModiPassword`** is spelled that way in the published package, and a
  workflow binds by property name.
- **`Delay` is a `double` here**, where the other packages use a `short` or an `int`.
- **Tables, rows and columns are numbered from 1**, matching Word's own object model
  rather than the 0-based convention.
- **`RowHeight` with no indexes means every row**, which is the only reading that makes
  an empty array useful rather than a silent no-op.
- **`EnumSelectBoolean` is `Select`=1, `True`=2, `False`=3**, so it cannot be cast to a
  Word boolean: `Select` would read as true and `False` as a nonzero truth. `WordService`
  maps it explicitly and leaves the run alone on `Select`.
- **The scope closes its document even when a child faults**, otherwise a failed run
  would leave a Word process behind.

For Excel:

- **There is no scope activity.** Each activity carries its own file path, sheet name
  and passwords, opens the workbook, works and closes it again, which is how the
  published package was shaped. An empty sheet name means the workbook's active sheet.
- **`ExecutionResult` is declared per activity, not on the base**, so the base cannot
  set it directly; `ReportResult` is the hook each activity overrides. `GetComment` is
  the one activity that reports something else and so overrides nothing.
- **Every formatting enum has a `Select` member meaning "leave this alone"**, which is
  what stops `FormatCells` overwriting formatting the workflow never mentioned.
- **Two misspellings are preserved**: the comment base is `BaseCommnet`, and
  `Enums.TextOrientationEumn` sits alongside a correctly spelled
  `Utilities.TextOrientationEnum`. Both are bound by workflows.
- **`InsertTableFormat` prefers `CustomStyle` over `TableFormatStyle`** when both are
  set, since a named style is the more specific instruction.
