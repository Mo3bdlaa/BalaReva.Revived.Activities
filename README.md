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

## Scope

The published BalaReva packages are closed-source; this repository does not decompile or
redistribute them. Reviving an activity here means reimplementing it against a supported
runtime and current dependencies, keeping the activity and property names compatible so
existing workflows keep binding.
