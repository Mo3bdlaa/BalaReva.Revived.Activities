# BalaReva Revived — Gzip Unzip Activities

Tar and gzip extraction for UiPath, reimplemented on .NET 8.

Type and property names match the published `BalaReva.ZipUnzipGz.Activities` 2.0.0, so
existing workflows keep binding — including the `Cls` suffix on `UnZipCls`.

**This package targets plain `net8.0`** and works in a Studio *Cross-platform* project,
which the original could not.

## Why it was worth reimplementing

The published package depended on `SharpZipLib` 1.2.0 **and shipped a copy of the assembly
in `lib/` alongside the dependency**, so a consumer could end up with the vulnerable
version regardless of what NuGet resolved. That version carried CVE-2021-32840 (HIGH) and
CVE-2021-32842, both path traversal on extraction, fixed in 1.3.3.

This package declares 1.4.2 and vendors nothing.

As with the zip package, the version bump is not the whole fix. Extraction reads the
archive entry by entry through `TarInputStream` rather than calling
`TarArchive.ExtractContents`, which is the method the advisories were written against
because it joins entry names to the destination itself. Every destination goes through
`SafeExtractPath` first — the same check the zip package uses, from the same source file.

## What it accepts

A `.gz` holding a single file, a `.tar.gz` or `.tgz`, and a plain `.tar`. The format is
decided by looking at the bytes rather than the file extension: gzip by its magic number,
tar by the `ustar` marker at offset 257 of its first header block.

## What is tested

Everything, against real archives on disk, on both CI jobs. Nothing here needs anything
installed.
