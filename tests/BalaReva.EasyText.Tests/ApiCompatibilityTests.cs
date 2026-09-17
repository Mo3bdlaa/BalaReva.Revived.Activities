using BalaReva.EasyText.Base;
using BalaReva.Revived.TestSupport;

namespace BalaReva.EasyText.Tests;

/// <summary>
/// Checks the reimplementation against the recorded surface of the published
/// BalaReva.EasyText.Activities 3.0.1.
/// </summary>
public class ApiCompatibilityTests
{
    private const string PackageId = "BalaReva.EasyText.Activities";

    private static readonly System.Reflection.Assembly Revived = typeof(BaseChildActivity).Assembly;

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
        // Guards against the whole suite passing vacuously on an empty list.
        Assert.Equal(13, PublishedSurface.Activities(PackageId).Count);
}
