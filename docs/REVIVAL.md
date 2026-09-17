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
| `BalaReva.EasyOutlook.Activities` | 20 | — | Not started |
| `BalaReva.Excel.Activities` | 39 | — | Not started |
| `BalaReva.Word.Activities` | 39 | — | Not started |
| `BalaReva.EasyPowerPoint.Activities` | 56 | — | Not started |
| `BalaReva.EasyExcel.Activities` | 77 | — | Not started |

EasyText went first because it is the only one of the eight with no Windows
dependency at all, which means its behaviour can be executed and asserted on any
agent rather than merely compiled.

### Why EasyImage and Printer have to stay on Windows

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

Both therefore target `net8.0-windows`. They are supported again, and usable from a
Studio *Windows* project, but they cannot be used from a *Cross-platform* project — and
no amount of retargeting would change that without breaking compatibility.

### How they are tested

The Windows-only packages cannot run on a Linux agent at all: `System.Drawing.Common`
throws there, and a `net8.0-windows` test host needs a Windows Desktop runtime that
does not exist on Linux. So CI has two jobs. The Linux job builds everything (which
works, via `EnableWindowsTargeting`) and runs the EasyText suite. The Windows job runs
all three suites.

Printer talks to the spooler through `IPrinterService`, and its tests register a
stand-in as a workflow extension. That is not only for portability: pausing a queue or
purging jobs on a build agent is destructive, and an agent may have no printers at all.
Everything worth testing — argument validation, which call is made with which
arguments, `ContinueOnError`, output mapping — sits on this side of that boundary. The
Windows implementation behind it is a thin translation, and is the part that a real
machine has to vouch for.

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
