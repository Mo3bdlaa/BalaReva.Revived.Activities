using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint.Utilities;
using BalaReva.Revived.TestSupport;

namespace BalaReva.EasyPowerPoint.Tests;

/// <summary>
/// Checks the reimplementation against the recorded surface of the published
/// BalaReva.EasyPowerPoint.Activities 11.0.0.
/// </summary>
public class ApiCompatibilityTests
{
    private const string PackageId = "BalaReva.EasyPowerPoint.Activities";

    private static readonly System.Reflection.Assembly Revived = typeof(BaseNativeChild).Assembly;

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
        Assert.Equal(61, PublishedSurface.Activities(PackageId).Count);

    [Fact]
    public void The_two_activities_in_the_odd_namespace_stayed_there()
    {
        // The published package put these two under BalaReva.Easy.PowerPoint - note the
        // extra dot - rather than BalaReva.EasyPowerPoint like the other 54. A workflow
        // binds by full type name, so moving them would be a breaking change.
        Assert.NotNull(Revived.GetType("BalaReva.Easy.PowerPoint.Scope.Slides.ExtractHyperLinks"));
        Assert.NotNull(Revived.GetType("BalaReva.Easy.PowerPoint.Scope.TableArea.GetRowItem"));
        Assert.Null(Revived.GetType("BalaReva.EasyPowerPoint.Scope.Slides.ExtractHyperLinks"));
    }

    [Fact]
    public void The_count_activities_kept_their_swapped_output_names()
    {
        // ImageShapeCount reports through ChartCount and TextShapeCount through
        // ImageCount. That is how the published package named them - apparently a
        // copy-and-paste slip - and both are bound by workflows.
        var images = Revived.GetType("BalaReva.EasyPowerPoint.Scope.Slides.ImageShapeCount")!;
        var text = Revived.GetType("BalaReva.EasyPowerPoint.Scope.Slides.TextShapeCount")!;

        Assert.NotNull(images.GetProperty("ChartCount"));
        Assert.Null(images.GetProperty("ImageCount"));
        Assert.NotNull(text.GetProperty("ImageCount"));
        Assert.Null(text.GetProperty("TextCount"));
    }

    [Fact]
    public void FindReplace_kept_its_underscore() =>
        Assert.NotNull(Revived.GetType("BalaReva.EasyPowerPoint.Scope.Slides.Find_Replace"));

    [Fact]
    public void The_transition_enum_kept_all_one_hundred_and_eighty_nine_effects() =>
        Assert.Equal(189, Enum.GetValues<Easy.PowerPoint.Utilities.EntryEffectEnum>().Length);

    [Fact]
    public void The_save_format_enum_kept_all_thirty_nine_formats() =>
        Assert.Equal(39, Enum.GetValues<Easy.PowerPoint.Utilities.SaveAsEnum>().Length);

    [Theory]
    [InlineData(TrueFalseNoneEnum.True, 0)]
    [InlineData(TrueFalseNoneEnum.False, 1)]
    [InlineData(TrueFalseNoneEnum.None, 2)]
    public void TrueFalseNone_values_match_the_published_package(TrueFalseNoneEnum value, int expected) =>
        // None is 2, not a negative sentinel, so it cannot be cast to a boolean.
        Assert.Equal(expected, (int)value);

    [Fact]
    public void Delay_is_a_double_here_as_it_is_in_Word()
    {
        var delay = typeof(BaseNativeChild).GetProperty("Delay")!;
        Assert.Equal(typeof(System.Activities.InArgument<double>), delay.PropertyType);
    }
}
