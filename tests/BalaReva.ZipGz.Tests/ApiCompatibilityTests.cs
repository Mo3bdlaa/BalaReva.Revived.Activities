using BalaReva.Revived.TestSupport;

namespace BalaReva.ZipGz.Tests;

/// <summary>
/// Checks the reimplementation against the recorded surface of the published
/// BalaReva.ZipUnzipGz.Activities 2.0.0.
/// </summary>
public class ApiCompatibilityTests
{
    private const string PackageId = "BalaReva.ZipUnzipGz.Activities";

    private static readonly System.Reflection.Assembly Revived = typeof(UnZipCls).Assembly;

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
        Assert.Single(PublishedSurface.Activities(PackageId));

    [Fact]
    public void The_single_activity_kept_its_Cls_suffix() =>
        Assert.NotNull(Revived.GetType("BalaReva.ZipGz.UnZipCls"));
}
