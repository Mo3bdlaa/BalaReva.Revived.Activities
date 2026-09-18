using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Tools;

/// <summary>Writes a slide table into a workbook.</summary>
/// <remarks>
/// Needs the Excel object model, which this package deliberately does not
/// depend on, so it copies the table to the clipboard and then reports that
/// the write is not implemented. See docs/REVIVAL.md.
/// </remarks>
[DisplayName("Export Table To Excel")]
[Description("Writes a slide table into a workbook.")]
public sealed class ExportTableToExcel : BaseTableNativeChild
{
    /// <summary>Full path of the workbook to write to.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Excel File")]
    [Description("Full path of the workbook to write to.")]
    public InArgument<string> ExcelFile { get; set; } = null!;

    /// <summary>Sheet to write into.</summary>
    [Category("Input")]
    [DisplayName("Sheet Name")]
    [Description("Sheet to write into.")]
    public InArgument<string> SheetName { get; set; } = null!;

    /// <summary>Cell to start writing at, for example A1.</summary>
    [Category("Input")]
    [DisplayName("Start Cell")]
    [Description("Cell to start writing at, for example A1.")]
    public InArgument<string> StartCell { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.ExportTableToExcel(
            Table(context),
            Require(context, ExcelFile, nameof(ExcelFile)),
            SheetName?.Get(context) ?? string.Empty,
            StartCell?.Get(context) ?? string.Empty);
}
