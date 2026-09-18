using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;
using BalaReva.Revived.TestSupport;

namespace BalaReva.EasyExcel.Tests;

/// <summary>
/// Checks the reimplementation against the recorded surface of the published
/// BalaReva.EasyExcel.Activities 32.0.0.
/// </summary>
public class ApiCompatibilityTests
{
    private const string PackageId = "BalaReva.EasyExcel.Activities";

    private static readonly System.Reflection.Assembly Revived = typeof(ExcelActivity).Assembly;

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
        Assert.Equal(77, PublishedSurface.Activities(PackageId).Count);

    [Fact]
    public void The_misspelled_enum_kept_its_name()
    {
        // TrueFaleNoneEnum, not TrueFalseNoneEnum. A workflow binds the type name, so the
        // typo is part of the contract.
        Assert.NotNull(Revived.GetType("BalaReva.EasyExcel.Utilities.TrueFaleNoneEnum"));
        Assert.Null(Revived.GetType("BalaReva.EasyExcel.Utilities.TrueFalseNoneEnum"));
    }

    [Fact]
    public void The_misspelled_hyperlink_argument_kept_its_name()
    {
        // DisplayTextOverwirte, not DisplayTextOverwrite.
        var type = Revived.GetType("BalaReva.EasyExcel.HyperLinks.InsertHyperlink")!;
        Assert.NotNull(type.GetProperty("DisplayTextOverwirte"));
        Assert.Null(type.GetProperty("DisplayTextOverwrite"));
    }

    [Fact]
    public void The_underscored_arguments_kept_their_underscores()
    {
        Assert.NotNull(Revived.GetType("BalaReva.EasyExcel.Settings.General")!
            .GetProperty("Adaptive_Menus"));
        Assert.NotNull(Revived.GetType("BalaReva.EasyExcel.WorkBook.TabColor")!
            .GetProperty("Tab_Color"));
        Assert.NotNull(Revived.GetType("BalaReva.EasyExcel.Sheets.CellFont")!
            .GetProperty("Font_Style"));
    }

    [Fact]
    public void Delay_is_a_short_here_where_Word_and_PowerPoint_use_a_double() =>
        Assert.Equal(
            typeof(System.Activities.InArgument<short>),
            typeof(ExcelActivity).GetProperty("Delay")!.PropertyType);

    [Fact]
    public void Both_chart_delete_all_activities_are_present()
    {
        // The published package shipped the same operation twice, under two names in two
        // namespaces. Both are part of the surface, so both are here.
        Assert.NotNull(Revived.GetType("BalaReva.EasyExcel.Charts.ChartDeleteAll"));
        Assert.NotNull(Revived.GetType("BalaReva.EasyExcel.Sheets.DeleteAllCharts"));
    }

    [Fact]
    public void The_file_format_enums_kept_all_fifty_four_formats()
    {
        Assert.Equal(54, Enum.GetValues<FileFormatEnum>().Length);
        Assert.Equal(54, Enum.GetValues<SaveEnum>().Length);
    }

    [Fact]
    public void The_paper_size_enum_kept_all_forty_three_sizes() =>
        Assert.Equal(43, Enum.GetValues<PaperSizeEnum>().Length);

    [Fact]
    public void SetBorder_still_binds_the_interop_line_style() =>
        // The one activity in the family with a COM type in its binding surface.
        Assert.Equal(
            typeof(Microsoft.Office.Interop.Excel.XlLineStyle),
            Revived.GetType("BalaReva.EasyExcel.Sheets.SetBorder")!
                .GetProperty("LineStyle")!.PropertyType);
}
