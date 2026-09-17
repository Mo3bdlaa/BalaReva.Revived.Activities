using System.Activities;
using BalaReva.EasyOutlook.Appointments;
using BalaReva.EasyOutlook.Contacts;
using BalaReva.EasyOutlook.EmailItem;
using BalaReva.EasyOutlook.Folders;
using BalaReva.EasyOutlook.Meetings;
using BalaReva.EasyOutlook.Notes;
using BalaReva.EasyOutlook.Utilities;

namespace BalaReva.EasyOutlook.Tests;

/// <summary>
/// Each activity runs as a real workflow against a stand-in Outlook session, covering
/// argument handling and output mapping without a mailbox.
/// </summary>
public class ActivityTests
{
    // ---------------------------------------------------------------- folders

    [Fact]
    public void GetFolders_reports_what_the_session_returns()
    {
        var service = new FakeOutlookService { Folders = ["Archive", "Projects"] };

        var outputs = Harness.Run(new GetFolders(), service);

        Assert.Equal(["Archive", "Projects"], (string[])outputs["FoldersList"]);
        Assert.Equal(["GetFolders"], service.Calls);
    }

    [Fact]
    public void GetSubFolders_passes_the_parent_through()
    {
        var service = new FakeOutlookService();

        Harness.Run(new GetSubFolders { Folder = new InArgument<string>("Archive") }, service);

        Assert.Equal(["GetSubFolders(Archive)"], service.Calls);
    }

    [Fact]
    public void CreateFolder_and_CreateSubFolder_pass_their_names_through()
    {
        var service = new FakeOutlookService();

        Harness.Run(new CreateFolder { FolderName = new InArgument<string>("Archive") }, service);
        Harness.Run(
            new CreateSubFolder
            {
                Folder = new InArgument<string>("Archive"),
                SubFolder = new InArgument<string>("2026"),
            },
            service);

        Assert.Equal(["CreateFolder(Archive)", "CreateSubFolder(Archive,2026)"], service.Calls);
    }

    [Fact]
    public void The_delete_activities_pass_their_names_through()
    {
        var service = new FakeOutlookService();

        Harness.Run(new DeleteFolder { FolderName = new InArgument<string>("Old") }, service);
        Harness.Run(
            new DeleteSubFolder
            {
                FolderName = new InArgument<string>("Archive"),
                SubFolderName = new InArgument<string>("2019"),
            },
            service);
        Harness.Run(new DeleteEmptyFolders(), service);
        Harness.Run(new DeleteEmptySubFolders { FolderName = new InArgument<string>("Archive") }, service);

        Assert.Equal(
            [
                "DeleteFolder(Old)",
                "DeleteSubFolder(Archive,2019)",
                "DeleteEmptyFolders",
                "DeleteEmptySubFolders(Archive)",
            ],
            service.Calls);
    }

    [Fact]
    public void The_rename_activities_pass_old_and_new_names_in_order()
    {
        var service = new FakeOutlookService();

        Harness.Run(
            new RenameFolder
            {
                ExistingFolderName = new InArgument<string>("Old"),
                NewFolderName = new InArgument<string>("New"),
            },
            service);
        Harness.Run(
            new RenameSubFolder
            {
                FolderName = new InArgument<string>("Archive"),
                ExistingSubFolderName = new InArgument<string>("2019"),
                NewSubFolderName = new InArgument<string>("2019-archived"),
            },
            service);

        Assert.Equal(
            ["RenameFolder(Old,New)", "RenameSubFolder(Archive,2019,2019-archived)"],
            service.Calls);
    }

    [Fact]
    public void The_SharePoint_activities_read_the_SharePoint_store()
    {
        var service = new FakeOutlookService { SharePointFolders = ["Team Site"] };

        var all = Harness.Run(new GetAllSharePointFolders(), service);
        Harness.Run(new GetAllSharePointSubFolders { Folder = new InArgument<string>("Team Site") }, service);

        Assert.Equal(["Team Site"], (string[])all["FoldersList"]);
        Assert.Equal(["GetSharePointFolders", "GetSharePointSubFolders(Team Site)"], service.Calls);
    }

    // ------------------------------------------------------------------ mail

    [Fact]
    public void GetMailMessages_forwards_every_selector()
    {
        var service = new FakeOutlookService
        {
            Mail = [new BalaReva.Outlook.EmailItem(), new BalaReva.Outlook.EmailItem()],
        };

        var outputs = Harness.Run(
            new GetMailMessages
            {
                Account = new InArgument<string>("work@example.com"),
                MailFolder = MailFolderEnum.SentMail,
                Filter = new InArgument<string>("[Subject] = 'Report'"),
            },
            service);

        Assert.Equal(2, ((BalaReva.Outlook.EmailItem[])outputs["MailItems"]).Length);
        Assert.Equal(
            ["GetMailMessages(work@example.com|SentMail||[Subject] = 'Report')"], service.Calls);
    }

    [Fact]
    public void GetMailMessages_defaults_to_the_inbox_and_the_default_account()
    {
        var service = new FakeOutlookService();

        Harness.Run(new GetMailMessages(), service);

        Assert.Equal(["GetMailMessages(|Inbox||)"], service.Calls);
    }

    [Fact]
    public void GetMailMessages_prefers_a_specific_folder_over_the_default_one()
    {
        var service = new FakeOutlookService();

        Harness.Run(
            new GetMailMessages
            {
                MailFolder = MailFolderEnum.Inbox,
                SpecificFolder = new InArgument<string>("Archive"),
            },
            service);

        // Both are passed on; choosing between them is the session's job.
        Assert.Equal(["GetMailMessages(|Inbox|Archive|)"], service.Calls);
    }

    // -------------------------------------------------------------- calendar

    [Fact]
    public void NewAppointment_gathers_its_arguments_into_the_entry()
    {
        var service = new FakeOutlookService();
        var start = new DateTime(2026, 4, 1, 9, 0, 0, DateTimeKind.Unspecified);

        var outputs = Harness.Run(
            new NewAppointment
            {
                Subject = new InArgument<string>("Review"),
                Body = new InArgument<string>("Quarterly review"),
                AppointmentLocation = new InArgument<string>("Room 3"),
                StartDate = new InArgument<DateTime>(start),
                EndDate = new InArgument<DateTime>(start.AddHours(1)),
                ReminderMinutes = new InArgument<int>(15),
                BusyStatus = BusyStatusEnum.OutOfOffice,
                Importance = ImportanceEnum.High,
            },
            service);

        Assert.True((bool)outputs["AddedResult"]);
        var entry = Assert.IsType<CalendarEntry>(service.LastEntry);
        Assert.Equal("Review", entry.Subject);
        Assert.Equal("Room 3", entry.Location);
        Assert.Equal(start, entry.Start);
        Assert.Equal(15, entry.ReminderMinutes);
        Assert.Equal(BusyStatusEnum.OutOfOffice, entry.BusyStatus);
        Assert.Equal(ImportanceEnum.High, entry.Importance);
    }

    [Fact]
    public void NewAppointment_rejects_an_end_before_the_start()
    {
        var service = new FakeOutlookService();
        var start = new DateTime(2026, 4, 1, 9, 0, 0, DateTimeKind.Unspecified);

        var activity = new NewAppointment
        {
            Subject = new InArgument<string>("Review"),
            StartDate = new InArgument<DateTime>(start),
            EndDate = new InArgument<DateTime>(start.AddHours(-1)),
        };

        Assert.ThrowsAny<Exception>(() => Harness.Run(activity, service));
        Assert.Empty(service.Calls);
    }

    [Fact]
    public void NewAppointment_lets_an_all_day_event_ignore_the_times()
    {
        var service = new FakeOutlookService();
        var day = new DateTime(2026, 4, 1, 17, 0, 0, DateTimeKind.Unspecified);

        Harness.Run(
            new NewAppointment
            {
                Subject = new InArgument<string>("Company holiday"),
                AllDayEvent = new InArgument<bool>(true),
                StartDate = new InArgument<DateTime>(day),
                // Earlier in the day than the start, which is fine for an all-day entry.
                EndDate = new InArgument<DateTime>(day.AddHours(-8)),
            },
            service);

        Assert.Equal(["AddAppointment(Company holiday)"], service.Calls);
    }

    [Fact]
    public void NewMeeting_carries_the_attendees()
    {
        var service = new FakeOutlookService();
        var start = new DateTime(2026, 4, 1, 9, 0, 0, DateTimeKind.Unspecified);

        var outputs = Harness.Run(
            new NewMeeting
            {
                Subject = new InArgument<string>("Kickoff"),
                StartDate = new InArgument<DateTime>(start),
                EndDate = new InArgument<DateTime>(start.AddHours(1)),
                RequiredAttendees = new InArgument<string>("a@example.com;b@example.com"),
                OptionalAttendees = new InArgument<string>("c@example.com"),
                ResponseRequested = new InArgument<bool>(true),
            },
            service);

        Assert.True((bool)outputs["AddedResult"]);
        var entry = Assert.IsType<MeetingEntry>(service.LastEntry);
        Assert.Equal("a@example.com;b@example.com", entry.RequiredAttendees);
        Assert.Equal("c@example.com", entry.OptionalAttendees);
        Assert.True(entry.ResponseRequested);
        Assert.Equal(["AddMeeting(Kickoff)"], service.Calls);
    }

    // -------------------------------------------------------- contacts, notes

    [Fact]
    public void GetContacts_passes_the_snapshots_through()
    {
        var service = new FakeOutlookService
        {
            Contacts = [new OutlookContact { FirstName = "Ada", LastName = "Lovelace" }],
        };

        var outputs = Harness.Run(new GetContacts(), service);

        var contacts = Assert.IsType<List<OutlookContact>>(outputs["ContactCollection"]);
        Assert.Equal("Ada", Assert.Single(contacts).FirstName);
    }

    [Fact]
    public void NewContact_passes_the_account_and_the_details()
    {
        var service = new FakeOutlookService();
        var contact = new OutlookNewContact { FirstName = "Ada", Email1Address = "ada@example.com" };

        var outputs = Harness.Run(
            new NewContact
            {
                Account = new InArgument<string>("work@example.com"),
                Contact = new InArgument<OutlookNewContact>(_ => contact),
            },
            service);

        Assert.True((bool)outputs["AddedResult"]);
        Assert.Same(contact, service.LastContact);
        Assert.Equal(["AddContact(work@example.com,Ada)"], service.Calls);
    }

    [Fact]
    public void GetAllNotes_and_NewNotes_read_and_write_note_bodies()
    {
        var service = new FakeOutlookService { Notes = ["first", "second"] };

        var read = Harness.Run(new GetAllNotes(), service);
        var written = Harness.Run(new NewNotes { NoteBody = new InArgument<string>("third") }, service);

        Assert.Equal(["first", "second"], (List<string>)read["NotesCollection"]);
        Assert.True((bool)written["AddedResult"]);
        Assert.Equal(["GetNotes", "AddNote(third)"], service.Calls);
    }

    // ----------------------------------------------------------------- macro

    [Fact]
    public void ExecuteMacro_forwards_the_name_and_arguments_and_returns_the_result()
    {
        var service = new FakeOutlookService { MacroResult = 42 };

        var outputs = Harness.Run(
            new ExecuteMacro
            {
                MacroName = new InArgument<string>("Project1.Module1.Run"),
                Arguments = new InArgument<object[]>(_ => new object[] { "a", 1 }),
            },
            service);

        Assert.Equal(42, outputs["Result"]);
        Assert.Equal(["ExecuteMacro(Project1.Module1.Run,[a,1])"], service.Calls);
    }

    [Fact]
    public void ExecuteMacro_copes_with_a_macro_that_returns_nothing()
    {
        var service = new FakeOutlookService { MacroResult = null };

        var outputs = Harness.Run(
            new ExecuteMacro { MacroName = new InArgument<string>("Project1.Module1.Run") }, service);

        Assert.Null(outputs["Result"]);
    }
}
