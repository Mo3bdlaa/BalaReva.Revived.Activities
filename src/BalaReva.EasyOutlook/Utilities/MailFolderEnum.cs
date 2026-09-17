namespace BalaReva.EasyOutlook.Utilities;

/// <summary>Which default mail folder to read from.</summary>
/// <remarks>
/// Values match <c>Microsoft.Office.Interop.Outlook.OlDefaultFolders</c> exactly, so
/// they are deliberately not consecutive. They must not be renumbered.
/// </remarks>
public enum MailFolderEnum
{
    /// <summary>Deleted Items.</summary>
    DeletedItems = 3,
    /// <summary>Outbox.</summary>
    Outbox = 4,
    /// <summary>Sent Items.</summary>
    SentMail = 5,
    /// <summary>Inbox.</summary>
    Inbox = 6,
    /// <summary>Drafts.</summary>
    Drafts = 16,
    /// <summary>Junk Email.</summary>
    Junk = 23,
}
