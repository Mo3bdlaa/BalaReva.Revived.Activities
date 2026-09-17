using System.Reflection;
using BalaReva.EasyOutlook.Utilities;
// Aliased as Interop, not Outlook: this package also declares a BalaReva.Outlook
// namespace, and inside BalaReva.EasyOutlook that sibling wins over a using alias.
using Interop = Microsoft.Office.Interop.Outlook;

namespace BalaReva.EasyOutlook;

/// <summary>
/// <see cref="IOutlookService"/> on top of the Outlook COM object model.
/// </summary>
/// <remarks>
/// Requires Windows with Outlook installed, so this type is never exercised by CI: no
/// build agent has Outlook. Everything here is therefore kept to a mechanical
/// translation, with the judgement living in the activities where a stand-in can
/// cover it.
///
/// A new <c>Application</c> is created per operation. Outlook's automation server is a
/// singleton, so this attaches to the running instance rather than starting a second
/// one. RCWs are left to the garbage collector; <c>Marshal.ReleaseComObject</c> is
/// discouraged on modern .NET and does more harm than good when a workflow may hold on
/// to a <c>MailItem</c> through <see cref="BalaReva.Outlook.EmailItem"/>.
/// </remarks>
public sealed class OutlookService : IOutlookService
{
    /// <summary>The shared instance used when a workflow registers no extension.</summary>
    public static OutlookService Instance { get; } = new();

    /// <summary>
    /// Name Outlook gives the store that holds SharePoint lists.
    /// </summary>
    /// <remarks>
    /// An assumption, and a localisable one: on a non-English Outlook this store is
    /// named differently and the SharePoint activities will not find it. See
    /// docs/REVIVAL.md.
    /// </remarks>
    private const string SharePointStoreName = "SharePoint Lists";

    private static Interop.Application App() => new();

    private static Interop.NameSpace Session() => App().GetNamespace("MAPI");

    // ---------------------------------------------------------------- macros

    /// <inheritdoc />
    public object? ExecuteMacro(string macroName, object[]? arguments)
    {
        // Outlook's object model has no documented Application.Run, unlike Excel and
        // Word. Invoking it late-bound is what reaches a VBA macro where the host
        // allows it; where it does not, the COM error is surfaced as-is rather than
        // being swallowed, because a silently ignored macro is worse than a failure.
        var application = App();
        var parameters = new object[(arguments?.Length ?? 0) + 1];
        parameters[0] = macroName;
        arguments?.CopyTo(parameters, 1);

        return application.GetType().InvokeMember(
            "Run", BindingFlags.InvokeMethod, null, application, parameters);
    }

    // -------------------------------------------------------------- calendar

    /// <inheritdoc />
    public bool AddAppointment(CalendarEntry entry)
    {
        var item = NewAppointmentItem(entry);
        item.Save();
        return true;
    }

    /// <inheritdoc />
    public bool AddMeeting(MeetingEntry entry)
    {
        var item = NewAppointmentItem(entry);
        item.MeetingStatus = Interop.OlMeetingStatus.olMeeting;
        item.RequiredAttendees = entry.RequiredAttendees;
        item.OptionalAttendees = entry.OptionalAttendees;
        item.ResponseRequested = entry.ResponseRequested;
        item.Send();
        return true;
    }

    private static Interop.AppointmentItem NewAppointmentItem(CalendarEntry entry)
    {
        ArgumentNullException.ThrowIfNull(entry);
        var item = (Interop.AppointmentItem)App().CreateItem(Interop.OlItemType.olAppointmentItem);

        item.Subject = entry.Subject;
        item.Body = entry.Body;
        item.Location = entry.Location;
        item.AllDayEvent = entry.AllDayEvent;
        item.Start = entry.Start;
        item.End = entry.End;
        item.BusyStatus = (Interop.OlBusyStatus)(int)entry.BusyStatus;
        item.Importance = (Interop.OlImportance)(int)entry.Importance;

        // Setting ReminderMinutesBeforeStart without ReminderSet has no effect, and a
        // zero or negative value is taken as "no reminder" rather than "remind now".
        item.ReminderSet = entry.ReminderMinutes > 0;
        if (entry.ReminderMinutes > 0) item.ReminderMinutesBeforeStart = entry.ReminderMinutes;

        foreach (var attachment in entry.Attachments)
        {
            if (!string.IsNullOrWhiteSpace(attachment)) item.Attachments.Add(attachment);
        }

        return item;
    }

    // -------------------------------------------------------------- contacts

    /// <inheritdoc />
    public List<OutlookContact> GetContacts()
    {
        var session = Session();
        var folder = session.GetDefaultFolder(Interop.OlDefaultFolders.olFolderContacts);
        var account = session.DefaultStore.DisplayName;

        var contacts = new List<OutlookContact>();
        foreach (var entry in folder.Items)
        {
            if (entry is Interop.ContactItem c) contacts.Add(Snapshot(c, account));
        }
        return contacts;
    }

    /// <inheritdoc />
    public bool AddContact(string account, OutlookNewContact contact)
    {
        ArgumentNullException.ThrowIfNull(contact);
        var item = (Interop.ContactItem)App().CreateItem(Interop.OlItemType.olContactItem);

        item.FirstName = contact.FirstName;
        item.MiddleName = contact.MiddleName;
        item.LastName = contact.LastName;
        item.NickName = contact.NickName;
        item.CompanyName = contact.CompanyName;
        item.Companies = contact.Companies;
        item.Email1Address = contact.Email1Address;
        item.Email2Address = contact.Email2Address;
        item.HomeTelephoneNumber = contact.HomeTelephoneNumber;
        item.Home2TelephoneNumber = contact.Home2TelephoneNumber;
        item.HomeFaxNumber = contact.HomeFaxNumber;
        item.MobileTelephoneNumber = contact.MobileTelephoneNumber;
        item.Business2TelephoneNumber = contact.Business2TelephoneNumber;
        item.CompanyMainTelephoneNumber = contact.CompanyMainTelephoneNumber;
        item.BusinessAddress = contact.BusinessAddress;
        item.BusinessAddressCity = contact.BusinessAddressCity;
        item.BusinessAddressCountry = contact.BusinessAddressCountry;
        item.BusinessAddressPostalCode = contact.BusinessAddressPostalCode;
        item.BusinessAddressPostOfficeBox = contact.BusinessAddressPostOfficeBox;
        item.BusinessAddressState = contact.BusinessAddressState;
        item.BusinessAddressStreet = contact.BusinessAddressStreet;
        item.HomeAddress = contact.HomeAddress;
        item.HomeAddressCity = contact.HomeAddressCity;
        item.HomeAddressCountry = contact.HomeAddressCountry;
        item.HomeAddressPostalCode = contact.HomeAddressPostalCode;
        item.HomeAddressPostOfficeBox = contact.HomeAddressPostOfficeBox;
        item.HomeAddressState = contact.HomeAddressState;
        item.HomeAddressStreet = contact.HomeAddressStreet;

        // Outlook rejects the sentinel DateTime.MinValue; leaving it unset is the way
        // to say "no birthday".
        if (contact.Birthday != default) item.Birthday = contact.Birthday;

        item.Save();
        return true;
    }

    private static OutlookContact Snapshot(Interop.ContactItem c, string account) => new()
    {
        Account = account,
        Birthday = SafeDate(() => c.Birthday),
        Business2TelephoneNumber = c.Business2TelephoneNumber ?? string.Empty,
        BusinessAddress = c.BusinessAddress ?? string.Empty,
        BusinessAddressCity = c.BusinessAddressCity ?? string.Empty,
        BusinessAddressCountry = c.BusinessAddressCountry ?? string.Empty,
        BusinessAddressPostalCode = c.BusinessAddressPostalCode ?? string.Empty,
        BusinessAddressPostOfficeBox = c.BusinessAddressPostOfficeBox ?? string.Empty,
        BusinessAddressState = c.BusinessAddressState ?? string.Empty,
        BusinessAddressStreet = c.BusinessAddressStreet ?? string.Empty,
        Companies = c.Companies ?? string.Empty,
        CompanyAndFullName = c.CompanyAndFullName ?? string.Empty,
        CompanyLastFirstNoSpace = c.CompanyLastFirstNoSpace ?? string.Empty,
        CompanyLastFirstSpaceOnly = c.CompanyLastFirstSpaceOnly ?? string.Empty,
        CompanyMainTelephoneNumber = c.CompanyMainTelephoneNumber ?? string.Empty,
        CompanyName = c.CompanyName ?? string.Empty,
        Email1Address = c.Email1Address ?? string.Empty,
        Email2Address = c.Email2Address ?? string.Empty,
        FirstName = c.FirstName ?? string.Empty,
        FullName = c.FullName ?? string.Empty,
        FullNameAndCompany = c.FullNameAndCompany ?? string.Empty,
        Home2TelephoneNumber = c.Home2TelephoneNumber ?? string.Empty,
        HomeAddress = c.HomeAddress ?? string.Empty,
        HomeAddressCity = c.HomeAddressCity ?? string.Empty,
        HomeAddressCountry = c.HomeAddressCountry ?? string.Empty,
        HomeAddressPostalCode = c.HomeAddressPostalCode ?? string.Empty,
        HomeAddressPostOfficeBox = c.HomeAddressPostOfficeBox ?? string.Empty,
        HomeAddressState = c.HomeAddressState ?? string.Empty,
        HomeAddressStreet = c.HomeAddressStreet ?? string.Empty,
        HomeFaxNumber = c.HomeFaxNumber ?? string.Empty,
        HomeTelephoneNumber = c.HomeTelephoneNumber ?? string.Empty,
        LastName = c.LastName ?? string.Empty,
        LastNameAndFirstName = c.LastNameAndFirstName ?? string.Empty,
        MiddleName = c.MiddleName ?? string.Empty,
        MobileTelephoneNumber = c.MobileTelephoneNumber ?? string.Empty,
        NickName = c.NickName ?? string.Empty,
    };

    /// <summary>
    /// Reads a date property that Outlook throws on when the contact has not set it.
    /// </summary>
    private static DateTime SafeDate(Func<DateTime> read)
    {
        try { return read(); }
        catch (System.Runtime.InteropServices.COMException) { return default; }
    }

    // ------------------------------------------------------------------ mail

    /// <inheritdoc />
    public BalaReva.Outlook.EmailItem[] GetMailMessages(
        string account, MailFolderEnum folder, string specificFolder, string filter)
    {
        var session = Session();
        var source = string.IsNullOrWhiteSpace(specificFolder)
            ? DefaultFolder(session, account, (Interop.OlDefaultFolders)(int)folder)
            : Child(Root(session, account), specificFolder);

        var items = source.Items;
        if (!string.IsNullOrWhiteSpace(filter)) items = items.Restrict(filter);

        var messages = new List<BalaReva.Outlook.EmailItem>();
        foreach (var entry in items)
        {
            if (entry is Interop.MailItem mail)
                messages.Add(new BalaReva.Outlook.EmailItem { OutlookEmailItem = mail });
        }
        return [.. messages];
    }

    // --------------------------------------------------------------- folders

    /// <inheritdoc />
    public string[] GetFolders() => Names(Root(Session(), string.Empty));

    /// <inheritdoc />
    public string[] GetSubFolders(string folder) => Names(Child(Root(Session(), string.Empty), folder));

    /// <inheritdoc />
    public void CreateFolder(string folderName) =>
        Root(Session(), string.Empty).Folders.Add(folderName);

    /// <inheritdoc />
    public void CreateSubFolder(string folder, string subFolder) =>
        Child(Root(Session(), string.Empty), folder).Folders.Add(subFolder);

    /// <inheritdoc />
    public void DeleteFolder(string folderName) =>
        Child(Root(Session(), string.Empty), folderName).Delete();

    /// <inheritdoc />
    public void DeleteSubFolder(string folderName, string subFolderName) =>
        Child(Child(Root(Session(), string.Empty), folderName), subFolderName).Delete();

    /// <inheritdoc />
    public void DeleteEmptyFolders() => DeleteEmptyChildren(Root(Session(), string.Empty));

    /// <inheritdoc />
    public void DeleteEmptySubFolders(string folderName) =>
        DeleteEmptyChildren(Child(Root(Session(), string.Empty), folderName));

    /// <inheritdoc />
    public void RenameFolder(string existingFolderName, string newFolderName) =>
        Child(Root(Session(), string.Empty), existingFolderName).Name = newFolderName;

    /// <inheritdoc />
    public void RenameSubFolder(string folderName, string existingSubFolderName, string newSubFolderName) =>
        Child(Child(Root(Session(), string.Empty), folderName), existingSubFolderName).Name =
            newSubFolderName;

    /// <inheritdoc />
    public string[] GetSharePointFolders() => Names(SharePointRoot(Session()));

    /// <inheritdoc />
    public string[] GetSharePointSubFolders(string folder) =>
        Names(Child(SharePointRoot(Session()), folder));

    // ----------------------------------------------------------------- notes

    /// <inheritdoc />
    public List<string> GetNotes()
    {
        var folder = Session().GetDefaultFolder(Interop.OlDefaultFolders.olFolderNotes);
        var notes = new List<string>();
        foreach (var entry in folder.Items)
        {
            if (entry is Interop.NoteItem note) notes.Add(note.Body ?? string.Empty);
        }
        return notes;
    }

    /// <inheritdoc />
    public bool AddNote(string noteBody)
    {
        var item = (Interop.NoteItem)App().CreateItem(Interop.OlItemType.olNoteItem);
        item.Body = noteBody;
        item.Save();
        return true;
    }

    // ------------------------------------------------------------- plumbing

    private static Interop.MAPIFolder Root(Interop.NameSpace session, string account)
    {
        if (string.IsNullOrWhiteSpace(account)) return session.DefaultStore.GetRootFolder();

        foreach (Interop.Store store in session.Stores)
        {
            if (string.Equals(store.DisplayName, account, StringComparison.OrdinalIgnoreCase))
                return store.GetRootFolder();
        }
        throw new ArgumentException($"No Outlook account named '{account}' is configured.", nameof(account));
    }

    private static Interop.MAPIFolder DefaultFolder(
        Interop.NameSpace session, string account, Interop.OlDefaultFolders folder)
    {
        if (string.IsNullOrWhiteSpace(account)) return session.GetDefaultFolder(folder);

        foreach (Interop.Store store in session.Stores)
        {
            if (string.Equals(store.DisplayName, account, StringComparison.OrdinalIgnoreCase))
                return store.GetDefaultFolder(folder);
        }
        throw new ArgumentException($"No Outlook account named '{account}' is configured.", nameof(account));
    }

    private static Interop.MAPIFolder SharePointRoot(Interop.NameSpace session)
    {
        foreach (Interop.Store store in session.Stores)
        {
            if (string.Equals(store.DisplayName, SharePointStoreName, StringComparison.OrdinalIgnoreCase))
                return store.GetRootFolder();
        }
        throw new InvalidOperationException(
            $"Outlook has no '{SharePointStoreName}' store, so there are no connected SharePoint lists.");
    }

    private static Interop.MAPIFolder Child(Interop.MAPIFolder parent, string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("A folder name is required.", nameof(name));

        foreach (Interop.MAPIFolder child in parent.Folders)
        {
            if (string.Equals(child.Name, name, StringComparison.OrdinalIgnoreCase)) return child;
        }
        throw new ArgumentException($"'{parent.Name}' has no folder named '{name}'.", nameof(name));
    }

    private static string[] Names(Interop.MAPIFolder parent)
    {
        var names = new List<string>();
        foreach (Interop.MAPIFolder child in parent.Folders) names.Add(child.Name);
        return [.. names];
    }

    /// <summary>Deletes the direct children of <paramref name="parent"/> that hold nothing.</summary>
    /// <remarks>
    /// Indexed backwards because Outlook's Folders collection is 1-based and reindexes
    /// on delete, so a forward loop would skip every other folder.
    /// </remarks>
    private static void DeleteEmptyChildren(Interop.MAPIFolder parent)
    {
        var folders = parent.Folders;
        for (var i = folders.Count; i >= 1; i--)
        {
            var child = folders[i];
            if (child.Items.Count == 0 && child.Folders.Count == 0) child.Delete();
        }
    }
}
