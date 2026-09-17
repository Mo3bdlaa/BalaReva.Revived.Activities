using BalaReva.Revived.TestSupport;

namespace BalaReva.EasyImage.Tests;

/// <summary>
/// Checks the reimplementation against the recorded surface of the published
/// BalaReva.EasyImage.Activities 3.0.1.
/// </summary>
public class ApiCompatibilityTests
{
    private const string PackageId = "BalaReva.EasyImage.Activities";

    private static readonly System.Reflection.Assembly Revived = typeof(BaseWork).Assembly;

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
        Assert.Equal(10, PublishedSurface.Activities(PackageId).Count);

    [Theory]
    [InlineData(EnumFileFormat.BMP, 1)]
    [InlineData(EnumFileFormat.EXIF, 2)]
    [InlineData(EnumFileFormat.EMF, 3)]
    [InlineData(EnumFileFormat.GIF, 4)]
    [InlineData(EnumFileFormat.JPEG, 5)]
    [InlineData(EnumFileFormat.PNG, 6)]
    [InlineData(EnumFileFormat.TIFF, 7)]
    [InlineData(EnumFileFormat.WMF, 8)]
    public void Format_enum_values_match_the_published_package(EnumFileFormat value, int expected) =>
        // A workflow can persist the underlying value, so these may not be renumbered.
        Assert.Equal(expected, (int)value);
}
