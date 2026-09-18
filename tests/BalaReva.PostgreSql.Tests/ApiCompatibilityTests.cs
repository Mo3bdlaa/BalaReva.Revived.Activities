using BalaReva.PostgreSQL;
using BalaReva.Revived.TestSupport;

namespace BalaReva.PostgreSql.Tests;

/// <summary>
/// Checks the reimplementation against the recorded surface of the published
/// BalaReva.PostgreSql.Activities 2021.0.0.
/// </summary>
public class ApiCompatibilityTests
{
    private const string PackageId = "BalaReva.PostgreSql.Activities";

    private static readonly System.Reflection.Assembly Revived = typeof(BaseData).Assembly;

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
        // Three command activities plus BaseData, which the published package declared
        // concrete rather than abstract.
        Assert.Equal(4, PublishedSurface.Activities(PackageId).Count);

    [Fact]
    public void The_namespace_keeps_its_shouting_SQL_while_the_assembly_does_not()
    {
        // Namespace BalaReva.PostgreSQL, assembly BalaReva.PostgreSql.Activities. Both
        // are as published, and a workflow binds the first.
        Assert.NotNull(Revived.GetType("BalaReva.PostgreSQL.ExecuteScalar"));
        Assert.Equal("BalaReva.PostgreSql.Activities", Revived.GetName().Name);
    }

    [Fact]
    public void BaseData_is_usable_on_its_own_as_the_published_package_had_it() =>
        Assert.False(typeof(BaseData).IsAbstract);

    [Fact]
    public void The_parameters_argument_keeps_its_Npgsql_type() =>
        // Unlike the Office interop assemblies, Npgsql resolves cleanly on .NET 8, so
        // there is no reason to weaken this to object.
        Assert.Equal(
            typeof(System.Activities.InArgument<Npgsql.NpgsqlParameter[]>),
            typeof(BaseData).GetProperty("Parameters")!.PropertyType);
}
