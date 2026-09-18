using BalaReva.Excel.Base;
using BalaReva.Excel.Enums;
using BalaReva.Revived.TestSupport;

namespace BalaReva.Excel.Tests;

/// <summary>
/// Checks the reimplementation against the recorded surface of the published
/// BalaReva.Excel.Activities 2021.1.0.
/// </summary>
public class ApiCompatibilityTests
{
    private const string PackageId = "BalaReva.Excel.Activities";

    private static readonly System.Reflection.Assembly Revived = typeof(ExcelCore).Assembly;

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
        Assert.Equal(44, PublishedSurface.Activities(PackageId).Count);

    [Fact]
    public void The_misspelled_comment_base_kept_its_name()
    {
        // BaseCommnet, not BaseComment. Misspelled in the published package, and part
        // of the binding surface, so it stays.
        Assert.NotNull(Revived.GetType("BalaReva.Excel.BaseCommnet"));
        Assert.Null(Revived.GetType("BalaReva.Excel.BaseComment"));
    }

    [Fact]
    public void The_misspelled_orientation_enum_kept_its_name()
    {
        // TextOrientationEumn, not TextOrientationEnum - and there is a correctly
        // spelled TextOrientationEnum in another namespace, so both have to exist.
        Assert.NotNull(Revived.GetType("BalaReva.Excel.Enums.TextOrientationEumn"));
        Assert.NotNull(Revived.GetType("BalaReva.Excel.Utilities.TextOrientationEnum"));
    }

    [Fact]
    public void The_table_format_enum_kept_all_sixty_one_styles() =>
        Assert.Equal(61, Enum.GetValues<TableFormatEnum>().Length);

    [Theory]
    [InlineData(AlignmentEnum.Right, -4152)]
    [InlineData(AlignmentEnum.Center, -4108)]
    [InlineData(AlignmentEnum.General, 1)]
    public void Alignment_values_match_Excel(AlignmentEnum value, int expected) =>
        Assert.Equal(expected, (int)value);

    [Theory]
    [InlineData(Charts.ColumnChartEnum.Column3D, -4100)]
    [InlineData(Charts.ColumnChartEnum.ColumnClustered, 51)]
    [InlineData(Charts.BarChartEnum.BarClustered, 57)]
    public void Chart_type_values_match_XlChartType(object value, int expected) =>
        Assert.Equal(expected, (int)value);

    [Fact]
    public void Every_format_enum_offers_a_Select_member() =>
        // Select means "leave this alone", which is what keeps FormatCells from
        // overwriting formatting the workflow never mentioned.
        Assert.All(
            new[]
            {
                typeof(FormatHorizontalEnum), typeof(FormatVerticleEnum),
                typeof(FormatTextControlEnum), typeof(FormatTextDirection),
                typeof(TextOrientationEumn),
            },
            type => Assert.Contains("Select", Enum.GetNames(type)));
}
