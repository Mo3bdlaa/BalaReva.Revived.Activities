# BalaReva Revived — Easy Outlook Activities

Outlook activities for UiPath, reimplemented on .NET 8.

Type and property names match the published `BalaReva.EasyOutlook.Activities` 3.0.0, so
existing workflows keep binding after the swap. A test suite checks the surface back
against the original on every build.

This package targets `net8.0-windows` and requires Outlook to be installed:
`BalaReva.Outlook.EmailItem` hands back a live
`Microsoft.Office.Interop.Outlook.MailItem`, so the COM type is part of the public API
and cannot be traded for anything portable.

Unlike the package it replaces, the Outlook interop assembly is a **declared NuGet
dependency** rather than a loose DLL copied into `lib/`. Vendored copies are invisible
to every manifest-based vulnerability scanner, which is one of the findings in
[docs/AUDIT.md](https://github.com/Mo3bdlaa/BalaReva.Revived.Activities/blob/main/docs/AUDIT.md).

Outlook access sits behind `IOutlookService`, so a workflow can register its own
implementation as an extension — useful for testing a process without a mailbox.

Activities: `GetMailMessages`, `GetContacts`, `NewContact`, `NewAppointment`,
`NewMeeting`, `GetAllNotes`, `NewNotes`, `ExecuteMacro`, and twelve folder activities
(`GetFolders`, `GetSubFolders`, `CreateFolder`, `CreateSubFolder`, `DeleteFolder`,
`DeleteSubFolder`, `DeleteEmptyFolders`, `DeleteEmptySubFolders`, `RenameFolder`,
`RenameSubFolder`, `GetAllSharePointFolders`, `GetAllSharePointSubFolders`).

**Note that no CI agent has Outlook installed**, so the COM layer of this package is not
covered by automated tests — only the activity layer above it is. See
[docs/REVIVAL.md](https://github.com/Mo3bdlaa/BalaReva.Revived.Activities/blob/main/docs/REVIVAL.md).
