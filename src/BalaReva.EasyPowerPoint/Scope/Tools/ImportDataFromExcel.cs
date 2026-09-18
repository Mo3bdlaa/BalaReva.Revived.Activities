using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Tools;

/// <summary>Fills a slide from a range in a workbook.</summary>
/// <remarks>
/// Needs the Excel object model, which this package deliberately does not
/// depend on. See docs/REVIVAL.md.
/// </remarks>
[DisplayName("Import Data From Excel")]
[Description("Fills a slide from a range in a workbook.")]
public sealed class ImportDataFromExcel : BaseSlideNativeChild
{
    /// <summary>Full path of the workbook to read.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Excel File")]
    [Description("Full path of the workbook to read.")]
    public InArgument<string> ExcelFile { get; set; } = null!;

    /// <summary>Sheet to read from.</summary>
    [Category("Input")]
    [DisplayName("Sheet Name")]
    [Description("Sheet to read from.")]
    public InArgument<string> SheetName { get; set; } = null!;

    /// <summary>Range to read, for example A1:D10.</summary>
    [Category("Input")]
    [DisplayName("Cell Range")]
    [Description("Range to read, for example A1:D10.")]
    public InArgument<string> CellRange { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.ImportDataFromExcel(
            SlideIndex.Get(context),
            Require(context, ExcelFile, nameof(ExcelFile)),
            SheetName?.Get(context) ?? string.Empty,
            CellRange?.Get(context) ?? string.Empty);
}
