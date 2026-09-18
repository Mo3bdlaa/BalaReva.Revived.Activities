using BalaReva.Revived.TestSupport;

namespace BalaReva.Word.Tests;

/// <summary>
/// Checks the reimplementation against the recorded surface of the published
/// BalaReva.Word.Activities 8.0.0.
/// </summary>
public class ApiCompatibilityTests
{
    private const string PackageId = "BalaReva.Word.Activities";

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
        Assert.Equal(42, PublishedSurface.Activities(PackageId).Count);

    [Fact]
    public void Delay_is_a_double_here_not_a_short()
    {
        // The other revived packages use short or int. Word's published base used a
        // double, and a workflow binds by argument type as well as by name.
        var delay = typeof(BaseNativeChild).GetProperty("Delay")!;
        Assert.Equal(typeof(System.Activities.InArgument<double>), delay.PropertyType);
    }

    [Fact]
    public void WordObject_kept_its_abbreviated_password_property()
    {
        // ModiPassword, not ModifyPassword. That is how the published package spelled
        // it, and a workflow binds by property name.
        Assert.NotNull(typeof(WordObject).GetProperty("ModiPassword"));
        Assert.Null(typeof(WordObject).GetProperty("ModifyPassword"));
    }

    [Fact]
    public void The_font_colour_enum_kept_all_sixty_Word_colours() =>
        Assert.Equal(60, Enum.GetValues<Utilities.EnumFontColor>().Length);

    [Theory]
    [InlineData(Utilities.EnumSaveAs.Document, 0)]
    [InlineData(Utilities.EnumSaveAs.PDF, 17)]
    [InlineData(Utilities.EnumSaveAs.StrictOpenXMLDocument, 24)]
    public void Save_format_values_match_WdSaveFormat(Utilities.EnumSaveAs value, int expected) =>
        Assert.Equal(expected, (int)value);

    [Theory]
    [InlineData(Utilities.EnumGoTo.Bookmark, -1)]
    [InlineData(Utilities.EnumGoTo.Page, 1)]
    [InlineData(Utilities.EnumGoTo.Line, 3)]
    public void GoTo_values_match_the_published_package(Utilities.EnumGoTo value, int expected) =>
        // Deliberately not consecutive: Page is 1 and Line is 3, with no 2.
        Assert.Equal(expected, (int)value);

    [Theory]
    [InlineData(Utilities.EnumSelectBoolean.Select, 1)]
    [InlineData(Utilities.EnumSelectBoolean.True, 2)]
    [InlineData(Utilities.EnumSelectBoolean.False, 3)]
    public void SelectBoolean_values_are_not_a_boolean(Utilities.EnumSelectBoolean value, int expected) =>
        // Select=1, True=2, False=3, so casting this to a Word boolean would read
        // Select as true and False as a nonzero truth. WordService maps it explicitly.
        Assert.Equal(expected, (int)value);
}
