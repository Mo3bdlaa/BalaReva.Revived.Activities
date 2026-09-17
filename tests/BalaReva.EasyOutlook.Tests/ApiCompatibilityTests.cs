using BalaReva.EasyOutlook.Utilities;
using BalaReva.Revived.TestSupport;

namespace BalaReva.EasyOutlook.Tests;

/// <summary>
/// Checks the reimplementation against the recorded surface of the published
/// BalaReva.EasyOutlook.Activities 3.0.0.
/// </summary>
public class ApiCompatibilityTests
{
    private const string PackageId = "BalaReva.EasyOutlook.Activities";

    private static readonly System.Reflection.Assembly Revived = typeof(BaseActivity).Assembly;

    public static TheoryData<string> Published() => PublishedSurface.ActivityNames(PackageId);

    [Theory]
    [MemberData(nameof(Published))]
    public void Every_published_activity_still_exists_under_the_same_full_name(string fullName) =>
        PublishedSurface.RequireType(Revived, fullName);

    [Theory]
    [MemberData(nameof(Published))]
    public void Every_published_property_still_exists_with_a_compatible_type(string fullName) =>
        PublishedSurface.RequireProperties(Revived, PackageId, fullName);

    [Fact]
    public void The_recorded_surface_was_actually_loaded() =>
        Assert.Equal(21, PublishedSurface.Activities(PackageId).Count);

    [Fact]
    public void EmailItem_kept_its_odd_namespace()
    {
        // The published package put this one type in BalaReva.Outlook rather than
        // BalaReva.EasyOutlook. It looks like an oversight, but a workflow binds by
        // full type name, so moving it would be a breaking change.
        Assert.NotNull(Revived.GetType("BalaReva.Outlook.EmailItem"));
        Assert.Null(Revived.GetType("BalaReva.EasyOutlook.EmailItem"));
    }

    [Theory]
    [InlineData(MailFolderEnum.DeletedItems, 3)]
    [InlineData(MailFolderEnum.Outbox, 4)]
    [InlineData(MailFolderEnum.SentMail, 5)]
    [InlineData(MailFolderEnum.Inbox, 6)]
    [InlineData(MailFolderEnum.Drafts, 16)]
    [InlineData(MailFolderEnum.Junk, 23)]
    public void Mail_folder_values_match_OlDefaultFolders(MailFolderEnum value, int expected) =>
        // Deliberately not consecutive: these are Outlook's own numbers.
        Assert.Equal(expected, (int)value);

    [Theory]
    [InlineData(BusyStatusEnum.Free, 0)]
    [InlineData(BusyStatusEnum.Tentative, 1)]
    [InlineData(BusyStatusEnum.Busy, 2)]
    [InlineData(BusyStatusEnum.OutOfOffice, 3)]
    [InlineData(BusyStatusEnum.WorkingElsewhere, 4)]
    public void Busy_status_values_match_OlBusyStatus(BusyStatusEnum value, int expected) =>
        Assert.Equal(expected, (int)value);

    [Theory]
    [InlineData(ImportanceEnum.Low, 0)]
    [InlineData(ImportanceEnum.Normal, 1)]
    [InlineData(ImportanceEnum.High, 2)]
    public void Importance_values_match_OlImportance(ImportanceEnum value, int expected) =>
        Assert.Equal(expected, (int)value);

    [Fact]
    public void The_contact_snapshots_still_carry_every_published_member()
    {
        Assert.Equal(36, typeof(OutlookContact).GetProperties().Length);
        Assert.Equal(29, typeof(OutlookNewContact).GetProperties().Length);
    }
}
