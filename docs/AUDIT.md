# BalaReva activity packages — dependency and compatibility audit

Generated 2026-09-17 from `https://gallery.uipath.com/api/v3/index.json`.
Regenerate with `python3 audit/balareva_audit.py`.

## Summary

- **28 packages** published under the BalaReva name on the UiPath Marketplace feed.
- **19 ship .NET Framework assemblies only** — unusable from a Studio *Windows* (.NET 6/8) or *Cross-platform* project; Legacy projects only.
- **8 carry .NET 6 assets**, so they load in Windows projects — but .NET 6 went out of support on 2024-11-12, so none of them target a supported runtime.
- **1 is mis-packaged**: a `lib/<tfm>/` folder whose assemblies target a different runtime family than the folder advertises.
- **3 dependencies carry advisories that cover the pinned version**, including two path-traversal classes in the archive handling.
- Newest release across the whole family: **2024-01-12**.


## Runtime compatibility

Each verdict comes from the assemblies themselves: the `TargetFrameworkAttribute` baked in by the compiler, cross-checked against whether the assembly binds `mscorlib` (.NET Framework) or `System.Runtime` (.NET Core / 5+). `lib/` with no framework folder means the assemblies sit directly in `lib/`, which NuGet treats as *any* framework even though the binaries are .NET Framework only.

| Package | Version | Published | `lib/` folders | Verdict |
|---|---|---|---|---|
| `BalaReva.EasyDataTable.Activities` | 5.0.0 | 2024-01-12 | `net6.0-windows7.0` | 🛑 Mis-packaged |
| `BalaReva.Access.Activities` | 2.0.2 | 2020-10-28 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.CSV.Activities` | 2019.1.0 | 2020-10-28 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.DataTable.Activities` | 2.0.2 | 2020-10-28 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.DirFileSize.Activities` | 2019.1.1 | 2020-10-28 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.EasyElasticsearch.Activities` | 1.0.0 | 2022-05-04 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.EasyLanguage.Activities` | 1.0.0 | 2020-10-28 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.Enterprise.EasyExcel.Activities` | 3.0.1 | 2022-03-16 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.Excel.Graph.Activities` | 2.0.1 | 2020-10-28 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.Externals.Activities` | 4.0.3 | 2020-10-28 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.FileProperties.Activities` | 1.0.0 | 2020-10-28 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.Informix.Activities` | 2019.2.0 | 2020-10-28 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.NumberFormatter.Activities` | 2019.0.0 | 2020-10-28 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.Pdf.Activities` | 2019.3.0 | 2020-10-28 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.PostgreSql.Activities` | 2021.0.0 | 2021-07-09 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.PowerPoint.Activities` | 3.0.2 | 2020-10-28 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.Web.Activities` | 3.0.0 | 2020-10-28 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.WindowsService.Activities` | 2.0.0 | 2020-10-28 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.ZipUnzip.Activities` | 2020.4.3 | 2021-10-12 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.ZipUnzipGz.Activities` | 2.0.0 | 2020-10-28 | `lib/` (no framework folder) | 🔴 Legacy only |
| `BalaReva.EasyExcel.Activities` | 32.0.0 | 2023-04-28 | `net461`, `net6.0-windows7.0` | 🟡 .NET 6 (EOL) |
| `BalaReva.EasyImage.Activities` | 3.0.1 | 2023-10-30 | `net461`, `net6.0-windows7.0` | 🟡 .NET 6 (EOL) |
| `BalaReva.EasyOutlook.Activities` | 3.0.0 | 2023-10-30 | `net461`, `net6.0-windows7.0` | 🟡 .NET 6 (EOL) |
| `BalaReva.EasyPowerPoint.Activities` | 11.0.0 | 2023-05-08 | `net461`, `net6.0-windows7.0` | 🟡 .NET 6 (EOL) |
| `BalaReva.EasyText.Activities` | 3.0.1 | 2023-10-30 | `net461`, `net6.0-windows7.0` | 🟡 .NET 6 (EOL) |
| `BalaReva.Excel.Activities` | 2021.1.0 | 2023-05-03 | `net461`, `net6.0-windows7.0` | 🟡 .NET 6 (EOL) |
| `BalaReva.Printer.Activities` | 2019.2.1 | 2023-10-30 | `net461`, `net6.0-windows7.0` | 🟡 .NET 6 (EOL) |
| `BalaReva.Word.Activities` | 8.0.0 | 2023-05-10 | `net461`, `net6.0-windows7.0` | 🟡 .NET 6 (EOL) |

### Mis-packaged assemblies

These load in Studio right up until the moment the designer needs them, because NuGet selects the folder by name and never verifies its contents.

- `BalaReva.EasyDataTable.Activities` 5.0.0 — `lib/net6.0-windows7.0/BalaReva.EasyDataTable.Design.dll`
  sits in a **net6.0-windows7.0** folder but targets **.NETFramework,Version=v4.6.1** (binds `mscorlib`).

## Dependency currency

| Dependency | Pinned | Latest on nuget.org | Advisories | Used by |
|---|---|---|---|---|
| `Npgsql` | 4.0.7 | 10.0.3 | CVE-2024-32655 | `BalaReva.PostgreSql.Activities` |
| `SharpCompress` | 0.26.0 | 1.0.0 | CVE-2026-44788, CVE-2021-39208 | `BalaReva.ZipUnzip.Activities` |
| `SharpZipLib` | 1.2.0 | 1.4.2 | CVE-2021-32840, CVE-2021-32842 | `BalaReva.ZipUnzipGz.Activities` |
| `DocumentFormat.OpenXml` | 2.20.0 | 3.5.1 | — | `BalaReva.EasyExcel.Activities`, `BalaReva.Excel.Activities` |
| `DotNetZip.Reduced` | 1.9.1.8 | 1.9.1.8 | — | `BalaReva.ZipUnzip.Activities` |
| `Elasticsearch.Net` | 6.0.0 | 7.17.5 | — | `BalaReva.EasyElasticsearch.Activities` |
| `Ionic.Zip` | 1.9.1.8 | 1.9.1.8 | — | `BalaReva.ZipUnzip.Activities` |
| `Microsoft.Office.Interop.Excel` | 15.0.4795.1000 | 16.0.18925.20022 | — | `BalaReva.CSV.Activities`, `BalaReva.Enterprise.EasyExcel.Activities` |
| `Microsoft.Office.Interop.PowerPoint` | 15.0.4420.1017 | 15.0.4420.1018 | — | `BalaReva.Enterprise.EasyExcel.Activities`, `BalaReva.PowerPoint.Activities` |
| `Microsoft.Practices.EnterpriseLibrary.Common.dll` | 3.1.0 | 3.1.0 | — | `BalaReva.Informix.Activities` |
| `Microsoft.Practices.EnterpriseLibrary.Data.dll` | 3.1.0 | 3.1.0 | — | `BalaReva.Informix.Activities` |
| `MicrosoftOfficeCore` | 15.0.0 | 15.0.0 | — | `BalaReva.Enterprise.EasyExcel.Activities` |
| `Newtonsoft.Json` | 13.0.1 | 13.0.4 | — | `BalaReva.EasyElasticsearch.Activities` |
| `iTextSharp` | 5.5.13.1 | 5.5.13.6 | — | `BalaReva.Pdf.Activities` |

## Known vulnerabilities

Only advisories whose affected range covers the **pinned** version are listed; each was queried against osv.dev per exact version rather than per package.

### CVE-2024-32655 — `Npgsql` 4.0.7  (HIGH)

Npgsql vulnerable to SQL Injection via Protocol Message Size Overflow

- Fixed in: 4.0.14, 4.1.13, 5.0.18, 6.0.11, 7.0.7, 8.0.3
- Reaches: `BalaReva.PostgreSql.Activities`

### CVE-2021-32840 — `SharpZipLib` 1.2.0  (HIGH)

Path Traversal in SharpZipLib

- Fixed in: 1.3.3
- Reaches: `BalaReva.ZipUnzipGz.Activities`

### CVE-2026-44788 — `SharpCompress` 0.26.0  (MODERATE)

SharpCompress has directory traversal via directory entries in WriteToDirectory (zip slip variant)

- Fixed in: 0.48.0
- Reaches: `BalaReva.ZipUnzip.Activities`

### CVE-2021-39208 — `SharpCompress` 0.26.0  (MODERATE)

Partial path traversal in sharpcompress

- Fixed in: 0.29
- Reaches: `BalaReva.ZipUnzip.Activities`

### CVE-2021-32842 — `SharpZipLib` 1.2.0  (MODERATE)

Path Traversal in SharpZipLib

- Fixed in: 1.3.3
- Reaches: `BalaReva.ZipUnzipGz.Activities`


## Caveats

- `DotNetZip.Reduced` and `Ionic.Zip` return clean from osv.dev because advisories are filed against the `DotNetZip` package id. They are the same codebase at the same version, so **CVE-2018-1002205 (zip-slip) applies to them too** — it is simply not reachable by an id lookup.
- Vendored copies of third-party libraries that were dropped into `lib/` instead of declared as dependencies are invisible to any manifest-based scanner, including this one's dependency table. They are listed in the package contents of `audit/data/audit.json`.
- Assembly-name evidence in `audit.json` comes from a heuristic scan of the `#Strings` heap; see `audit/pe_metadata.py`. The runtime verdicts above do not rely on it.
