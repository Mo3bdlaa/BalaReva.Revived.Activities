using BalaReva.EasyOutlook.Utilities;

namespace BalaReva.EasyOutlook.Tests;

/// <summary>
/// A stand-in Outlook session that records what it was asked to do.
/// </summary>
/// <remarks>
/// No CI agent has Outlook installed, so this is the only way any of these activities
/// get exercised. It covers the half that is worth covering: argument handling, which
/// service call is made with which values, and how results are mapped back onto output
/// arguments. The COM half is not reachable from a test at all.
/// </remarks>
public sealed class FakeOutlookService : IOutlookService
{
    public List<string> Calls { get; } = [];

    public object? MacroResult { get; set; }

    public List<OutlookContact> Contacts { get; set; } = [];

    public List<string> Notes { get; set; } = [];

    public string[] Folders { get; set; } = ["Archive", "Projects"];

    public string[] SharePointFolders { get; set; } = ["Team Site"];

    public BalaReva.Outlook.EmailItem[] Mail { get; set; } = [];

    public CalendarEntry? LastEntry { get; private set; }

    public OutlookNewContact? LastContact { get; private set; }

    private T Record<T>(string call, T result)
    {
        Calls.Add(call);
        return result;
    }

    public object? ExecuteMacro(string macroName, object[]? arguments) =>
        Record($"ExecuteMacro({macroName},[{string.Join(",", arguments ?? [])}])", MacroResult);

    public bool AddAppointment(CalendarEntry entry)
    {
        LastEntry = entry;
        return Record($"AddAppointment({entry.Subject})", true);
    }

    public bool AddMeeting(MeetingEntry entry)
    {
        LastEntry = entry;
        return Record($"AddMeeting({entry.Subject})", true);
    }

    public List<OutlookContact> GetContacts() => Record("GetContacts", Contacts);

    public bool AddContact(string account, OutlookNewContact contact)
    {
        LastContact = contact;
        return Record($"AddContact({account},{contact.FirstName})", true);
    }

    public BalaReva.Outlook.EmailItem[] GetMailMessages(
        string account, MailFolderEnum folder, string specificFolder, string filter) =>
        Record($"GetMailMessages({account}|{folder}|{specificFolder}|{filter})", Mail);

    public string[] GetFolders() => Record("GetFolders", Folders);

    public string[] GetSubFolders(string folder) => Record($"GetSubFolders({folder})", Folders);

    public void CreateFolder(string folderName) => Record($"CreateFolder({folderName})", 0);

    public void CreateSubFolder(string folder, string subFolder) =>
        Record($"CreateSubFolder({folder},{subFolder})", 0);

    public void DeleteFolder(string folderName) => Record($"DeleteFolder({folderName})", 0);

    public void DeleteSubFolder(string folderName, string subFolderName) =>
        Record($"DeleteSubFolder({folderName},{subFolderName})", 0);

    public void DeleteEmptyFolders() => Record("DeleteEmptyFolders", 0);

    public void DeleteEmptySubFolders(string folderName) =>
        Record($"DeleteEmptySubFolders({folderName})", 0);

    public void RenameFolder(string existingFolderName, string newFolderName) =>
        Record($"RenameFolder({existingFolderName},{newFolderName})", 0);

    public void RenameSubFolder(string folderName, string existingSubFolderName, string newSubFolderName) =>
        Record($"RenameSubFolder({folderName},{existingSubFolderName},{newSubFolderName})", 0);

    public string[] GetSharePointFolders() => Record("GetSharePointFolders", SharePointFolders);

    public string[] GetSharePointSubFolders(string folder) =>
        Record($"GetSharePointSubFolders({folder})", SharePointFolders);

    public List<string> GetNotes() => Record("GetNotes", Notes);

    public bool AddNote(string noteBody) => Record($"AddNote({noteBody})", true);
}

/// <summary>Runs one activity as a workflow root with a stand-in Outlook session.</summary>
internal static class Harness
{
    public static IDictionary<string, object> Run(BaseActivity activity, FakeOutlookService service)
    {
        activity.Delay ??= new System.Activities.InArgument<short>(0);

        var invoker = new System.Activities.WorkflowInvoker(activity);
        invoker.Extensions.Add(service);
        return invoker.Invoke();
    }
}
