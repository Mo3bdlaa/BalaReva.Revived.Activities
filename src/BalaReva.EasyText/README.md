# BalaReva Revived — Easy Text Activities

Text file activities for UiPath, reimplemented on .NET 8.

Type and property names match the published `BalaReva.EasyText.Activities` 3.0.1, so
existing workflows keep binding after the swap. A test suite checks the surface back
against the original on every build.

Unlike the package it replaces, this one targets plain `net8.0`: it works in Studio
*Windows* projects and in *Cross-platform* projects, and it preserves each file's
encoding, line ending and trailing newline on write.

Activities: `TextScope`, `ReadAll`, `ReadAllArray`, `ReadSpecificLine`, `LineCount`,
`FindText`, `FindLineIndex`, `FindReplaceAllText`, `FindReplaceAt`, `InsertLineAt`,
`DeleteAt`, `RemoveEmptyLine`.

Line numbers are 1-based; character offsets are 0-based. See
[docs/REVIVAL.md](https://github.com/Mo3bdlaa/BalaReva.Revived.Activities/blob/main/docs/REVIVAL.md).
