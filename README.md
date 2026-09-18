# BalaReva.Revived.Activities

The [BalaReva](https://marketplace.uipath.com/) activity packages are a widely used
family of community UiPath activities — Excel, Word, PowerPoint, PDF, Zip, Outlook,
PostgreSQL and more. They have effectively stopped receiving updates: **nothing in the
family has shipped since January 2024**, and most of it since October 2020.

This repository exists to work out exactly what has rotted and to revive the parts that
are worth reviving.

## Where things stand

Run `python3 audit/balareva_audit.py` for the current picture. As of the last run:

| | |
|---|---|
| Packages published under the BalaReva name | 28 |
| Ship **.NET Framework assemblies only** (Legacy projects only) | 19 |
| Ship .NET 6 assets — but .NET 6 left support 2024-11-12 | 8 |
| Mis-packaged (folder framework ≠ assembly framework) | 1 |
| Dependencies with advisories covering the pinned version | 3 |

The headline problem is not any single CVE — it is that **two thirds of the family cannot
be referenced from a modern UiPath project at all.** UiPath Studio's *Windows* (.NET 6/8)
and *Cross-platform* project types cannot consume a package that only carries
`.NETFramework` binaries, so those 19 packages are reachable only from Legacy projects.

The full evidence, per package and per dependency, is in **[docs/AUDIT.md](docs/AUDIT.md)**.

### Worth knowing about

- **`BalaReva.EasyDataTable.Activities` 5.0.0 is mis-packaged.** Its design-time assembly
  sits in `lib/net6.0-windows7.0/` but is compiled against .NET Framework 4.6.1. NuGet
  picks folders by name and never checks what is inside them, so this resolves cleanly and
  then fails when the designer actually needs the assembly.
- **The archive activities carry path-traversal bugs.** `BalaReva.ZipUnzip.Activities`
  pins SharpCompress 0.26.0 (CVE-2021-39208, CVE-2026-44788) alongside DotNetZip/Ionic.Zip
  1.9.1.8 (CVE-2018-1002205), and `BalaReva.ZipUnzipGz.Activities` pins SharpZipLib 1.2.0
  (CVE-2021-32840, CVE-2021-32842). All are zip-slip variants: a crafted archive writes
  outside the extraction directory.
- **`BalaReva.PostgreSql.Activities` pins Npgsql 4.0.7**, inside the affected range for
  CVE-2024-32655 (SQL injection via protocol message size overflow), fixed in 4.0.14.
- **Several packages rely on Office COM interop**, which pins them to Windows with a
  matching Office install and rules out cross-platform use regardless of retargeting.

## The audit tool

`audit/balareva_audit.py` is self-contained — Python 3, standard library only, no
`dotnet` required. It reads the UiPath Marketplace gallery feed, downloads each package,
and reports on it from evidence rather than from metadata:

- Runtime targets come from each assembly's `TargetFrameworkAttribute`, read straight out
  of the CLI metadata blob heap by `audit/pe_metadata.py`, and are cross-checked against
  whether the assembly binds `mscorlib` or `System.Runtime`. That is how the mis-packaged
  assembly above was found: the `lib/` folder name and the binary disagree.
- Vulnerabilities are queried from [osv.dev](https://osv.dev) **per pinned version**, so
  only advisories whose affected range actually covers the version in use are reported.

```
python3 audit/balareva_audit.py              # refresh from the network, regenerate both outputs
python3 audit/balareva_audit.py --offline    # rebuild the report from the local cache
```

Outputs are `docs/AUDIT.md` (readable) and `audit/data/audit.json` (machine-readable,
committed so changes between runs show up in a diff).

## Reviving them

The published packages are closed-source, so reviving one means reimplementing it
against a supported runtime rather than retargeting a fork. What gets preserved is the
binding surface: a `.xaml` workflow refers to an activity by type full name and to its
inputs and outputs by property name, so those have to survive the swap exactly.

`tools/ApiSurface` records the published surface from the assemblies' metadata tables —
names and signatures only, never method bodies — and the reimplementation is checked
back against that recording on every build. Rename a property and the build fails
instead of someone's workflow.

The current pass targets the eight packages that ship .NET 6 assets, since .NET 6 is out
of support. Between them they hold **262 concrete activities**:

| Package | Activities | Target | Status |
|---|---:|---|---|
| `BalaReva.EasyText.Activities` | 12 | `net8.0` | ✅ Reimplemented |
| `BalaReva.EasyImage.Activities` | 9 | `net8.0-windows` | ✅ Reimplemented |
| `BalaReva.Printer.Activities` | 10 | `net8.0-windows` | ✅ Reimplemented |
| `BalaReva.EasyOutlook.Activities` | 20 | `net8.0-windows` | ✅ Reimplemented |
| `BalaReva.Excel.Activities` | 39 | — | Not started |
| `BalaReva.Word.Activities` | 39 | `net8.0-windows` | ✅ Reimplemented |
| `BalaReva.EasyPowerPoint.Activities` | 56 | — | Not started |
| `BalaReva.EasyExcel.Activities` | 77 | — | Not started |

EasyText moved to plain `net8.0`, so it now works in Cross-platform projects too —
something the original could not do. The other three cannot follow it, and not for want
of effort: their **binding surfaces are made of Windows types**. EasyImage takes a
`Font`, a `Color`, a `Point` and a `RotateFlipType` as arguments; Printer's
`PrinterStatus` is `System.Printing.PrintQueue` member for member; EasyOutlook hands
back a live `Microsoft.Office.Interop.Outlook.MailItem`. Substituting a portable library
would break the very workflows these packages exist to keep working, so all three target
`net8.0-windows`.

EasyOutlook also carries a caveat worth reading before trusting it: **no build agent has
Outlook installed, so its COM layer is not covered by any automated test.** The activity
layer above it is, through a stand-in session.

That splits CI in two. Everything builds on Linux via `EnableWindowsTargeting`, but
`System.Drawing` throws there and a `net8.0-windows` test host needs a runtime Linux
does not have — so the Linux job runs the EasyText suite and the Windows job runs all
four.

[docs/REVIVAL.md](docs/REVIVAL.md) covers the approach, and records the behavioural
decisions that metadata could not settle — line numbering chief among them.

## Building

```
dotnet build BalaReva.Revived.sln -c Release
dotnet test  BalaReva.Revived.sln -c Release
dotnet pack  src/BalaReva.EasyText/BalaReva.EasyText.csproj -c Release -o artifacts
python3 audit/verify_package_layout.py 'artifacts/*.nupkg'
```

That last step is the audit's own PE reader pointed at our output: it fails a package
whose `lib/<tfm>/` folder disagrees with what its assemblies target, which is exactly
the defect found in `BalaReva.EasyDataTable.Activities` 5.0.0. CI runs it on every
build.
