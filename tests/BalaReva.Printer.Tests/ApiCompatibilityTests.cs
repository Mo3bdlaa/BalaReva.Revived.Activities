using BalaReva.Printer.Enums;
using BalaReva.Revived.TestSupport;

namespace BalaReva.Printer.Tests;

/// <summary>
/// Checks the reimplementation against the recorded surface of the published
/// BalaReva.Printer.Activities 2019.2.1.
/// </summary>
public class ApiCompatibilityTests
{
    private const string PackageId = "BalaReva.Printer.Activities";

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
        Assert.Equal(11, PublishedSurface.Activities(PackageId).Count);

    [Theory]
    [InlineData(AccessRightsEnum.None, 0)]
    [InlineData(AccessRightsEnum.AdministrateServer, 983041)]
    [InlineData(AccessRightsEnum.EnumerateServer, 131074)]
    [InlineData(AccessRightsEnum.UsePrinter, 131080)]
    [InlineData(AccessRightsEnum.AdministratePrinter, 983052)]
    public void Access_rights_values_match_PrintSystemDesiredAccess(AccessRightsEnum value, int expected) =>
        Assert.Equal(expected, (int)value);

    [Theory]
    [InlineData(CommandEnum.Pause, 1)]
    [InlineData(CommandEnum.RemoveAllJobs, 2)]
    [InlineData(CommandEnum.Refresh, 3)]
    [InlineData(CommandEnum.Resume, 4)]
    public void Command_values_match_the_published_package(CommandEnum value, int expected) =>
        Assert.Equal(expected, (int)value);

    [Fact]
    public void PrinterStatus_still_carries_every_published_member()
    {
        // 38 booleans plus FullName, each mirroring a System.Printing.PrintQueue member.
        var properties = typeof(PrinterStatus).GetProperties();
        Assert.Equal(39, properties.Length);
        Assert.Equal(38, properties.Count(p => p.PropertyType == typeof(bool)));
    }
}
