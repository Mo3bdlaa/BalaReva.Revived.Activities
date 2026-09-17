using BalaReva.EasyOutlook.Utilities;

namespace BalaReva.EasyOutlook;

/// <summary>
/// The Outlook operations the activities need.
/// </summary>
/// <remarks>
/// The activities talk to this rather than to the COM object model directly, so that
/// argument handling and output mapping can be tested with a stand-in. That matters
/// more here than anywhere else in this repository: no CI agent has Outlook installed,
/// so <see cref="OutlookService"/> itself cannot be exercised automatically at all.
/// Keeping it thin and mechanical is the only defence it gets.
///
/// A workflow can supply its own through <c>WorkflowInvoker.Extensions</c>; with none
/// registered the activities use the real one.
/// </remarks>
public interface IOutlookService
{
    /// <summary>Runs a VBA macro and returns its result.</summary>
    object? ExecuteMacro(string macroName, object[]? arguments);

    // Calendar

    /// <summary>Adds an appointment. Returns true when it was saved.</summary>
    bool AddAppointment(CalendarEntry entry);

    /// <summary>Sends a meeting invitation. Returns true when it was sent.</summary>
    bool AddMeeting(MeetingEntry entry);

    // Contacts

    /// <summary>Every contact in the default contacts folder.</summary>
    List<OutlookContact> GetContacts();

    /// <summary>Adds a contact. Returns true when it was saved.</summary>
    bool AddContact(string account, OutlookNewContact contact);

    // Mail

    /// <summary>
    /// Messages from a folder, optionally narrowed by a DASL or Jet restriction.
    /// </summary>
    BalaReva.Outlook.EmailItem[] GetMailMessages(
        string account, MailFolderEnum folder, string specificFolder, string filter);

    // Folders

    /// <summary>Names of the folders directly under the mailbox root.</summary>
    string[] GetFolders();

    /// <summary>Names of the folders directly under <paramref name="folder"/>.</summary>
    string[] GetSubFolders(string folder);

    /// <summary>Creates a folder under the mailbox root.</summary>
    void CreateFolder(string folderName);

    /// <summary>Creates a folder under an existing one.</summary>
    void CreateSubFolder(string folder, string subFolder);

    /// <summary>Deletes a folder under the mailbox root, and everything in it.</summary>
    void DeleteFolder(string folderName);

    /// <summary>Deletes a folder nested under another, and everything in it.</summary>
    void DeleteSubFolder(string folderName, string subFolderName);

    /// <summary>Deletes every empty folder under the mailbox root.</summary>
    void DeleteEmptyFolders();

    /// <summary>Deletes every empty folder under <paramref name="folderName"/>.</summary>
    void DeleteEmptySubFolders(string folderName);

    /// <summary>Renames a folder under the mailbox root.</summary>
    void RenameFolder(string existingFolderName, string newFolderName);

    /// <summary>Renames a folder nested under another.</summary>
    void RenameSubFolder(string folderName, string existingSubFolderName, string newSubFolderName);

    /// <summary>Names of the SharePoint list folders connected to Outlook.</summary>
    string[] GetSharePointFolders();

    /// <summary>Names of the folders under a connected SharePoint list folder.</summary>
    string[] GetSharePointSubFolders(string folder);

    // Notes

    /// <summary>The body text of every note in the default notes folder.</summary>
    List<string> GetNotes();

    /// <summary>Adds a note. Returns true when it was saved.</summary>
    bool AddNote(string noteBody);
}
