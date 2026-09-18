using BalaReva.Revived.TestSupport;

namespace BalaReva.Zip.Tests;

/// <summary>
/// Checks the reimplementation against the recorded surface of the published
/// BalaReva.ZipUnzip.Activities 2020.4.3.
/// </summary>
public class ApiCompatibilityTests
{
    private const string PackageId = "BalaReva.ZipUnzip.Activities";

    private static readonly System.Reflection.Assembly Revived = typeof(EnumExtractType).Assembly;

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
        Assert.Equal(2, PublishedSurface.Activities(PackageId).Count);

    [Fact]
    public void The_hungarian_prefix_on_the_archive_argument_stayed()
    {
        // strZipFile, not ZipFile. A workflow binds by property name, so the prefix is
        // part of the contract.
        var type = Revived.GetType("BalaReva.ZipUnzip.UnZipFile")!;
        Assert.NotNull(type.GetProperty("strZipFile"));
        Assert.Null(type.GetProperty("ZipFile"));
    }

    [Fact]
    public void The_compression_activity_kept_its_Cls_suffix() =>
        Assert.NotNull(Revived.GetType("BalaReva.ZipUnzip.ZipFilesCls"));

    [Fact]
    public void The_two_activities_and_their_enum_live_in_different_namespaces()
    {
        // The activities are under BalaReva.ZipUnzip; the enum one of them binds is under
        // BalaReva.Zip. Both are as published.
        Assert.NotNull(Revived.GetType("BalaReva.Zip.EnumExtractType"));
        Assert.NotNull(Revived.GetType("BalaReva.ZipUnzip.UnZipFile"));
    }

    [Fact]
    public void The_extract_type_enum_kept_both_members_and_their_numbers()
    {
        Assert.Equal(0, (int)EnumExtractType.Standard);
        Assert.Equal(1, (int)EnumExtractType.UniExtract);
        Assert.Equal(2, Enum.GetValues<EnumExtractType>().Length);
    }
}
