# BalaReva Revived — Zip and Unzip Activities

Archive activities for UiPath, reimplemented on .NET 8.

Type and property names match the published `BalaReva.ZipUnzip.Activities` 2020.4.3, so
existing workflows keep binding after the swap — including `strZipFile`, which keeps its
Hungarian prefix because a workflow binds by property name.

**This package targets plain `net8.0`.** Nothing in it touches Windows, so unlike the
original it works in a Studio *Cross-platform* project as well as a *Windows* one.

## Why it was worth reimplementing

The published package depended on three archive libraries, and every one of them carried
the same class of bug:

| Dependency | Pinned | Problem |
|---|---|---|
| `SharpCompress` | 0.26.0 | CVE-2021-39208, CVE-2026-44788 — path traversal on extraction |
| `SharpZipLib` | 1.2.0 | CVE-2021-32840, CVE-2021-32842 — path traversal on extraction |
| `DotNetZip.Reduced` / `Ionic.Zip` | 1.9.1.8 | Unmaintained since 2011; CVE-2018-1002205 has no fix |

All of them are **zip slip**: an archive entry named `../../etc/passwd`, or an absolute
path, joined to the destination folder without checking where it lands.

Bumping the versions is not the whole fix. This package extracts entry by entry and
resolves each destination itself, refusing anything outside the extraction folder, before
a single byte is written. That means an advisory against a future version of a dependency
is not automatically an advisory against these activities. `SafeExtractPath` is the check,
it is shared with the gzip package so there is one copy of it, and it is tested directly
against traversal, absolute, drive-qualified and UNC entry names.

DotNetZip is gone: it is unmaintained, and it cannot write an encrypted archive on .NET 8.
Password-protected zips are written and read through SharpZipLib 1.4.2 instead.

## What is tested

Everything. This is the first package in the repository where that is true: there is no
Office install and no database to stand in for, so `ArchiveService` is exercised against
real archives written and read on disk, on both CI jobs.
