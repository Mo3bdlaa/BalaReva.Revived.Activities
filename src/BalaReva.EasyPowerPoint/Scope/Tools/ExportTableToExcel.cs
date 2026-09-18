using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Tools;

/// <summary>Writes a slide table into a workbook.</summary>
/// <remarks>
/// Needs the Excel object model, which this package deliberately does not depend on,
/// so it copies the table to the clipboard and then says plainly that the write is not
/// implemented rather than failing silently. See docs/REVIVAL.md.
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
    {
        var excelFile = Require(context, ExcelFile, nameof(ExcelFile));

        // Put the table on the clipboard so the caller still has a route to the data.
        presentation.TableCopyToClipboard(Table(context));

        throw new NotSupportedException(
            $"Exporting a table to '{excelFile}' is not implemented. Writing the workbook needs "
            + "the Excel object model, which this package deliberately does not depend on. The "
            + "table has been copied to the clipboard; use BalaReva.Revived.Excel.Activities to "
            + "write it. See docs/REVIVAL.md.");
    }
}
