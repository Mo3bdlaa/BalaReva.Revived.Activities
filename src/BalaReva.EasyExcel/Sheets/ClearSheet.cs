using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Clears the chosen parts of a range.</summary>
[DisplayName("Clear Sheet")]
[Description("Clears the chosen parts of a range.")]
public sealed class ClearSheet : BaseActivity
{
    /// <summary>Range to clear. Empty means the whole sheet.</summary>
    [Category("Input")]
    [DisplayName("Cell Range")]
    [Description("Range to clear. Empty means the whole sheet.")]
    public InArgument<string> CellRange { get; set; } = null!;

    /// <summary>Clear everything.</summary>
    [Category("Input")]
    [DisplayName("Clear All")]
    [Description("Clear everything.")]
    public InArgument<bool> ClearAll { get; set; } = null!;

    /// <summary>Clear the threaded comments.</summary>
    [Category("Input")]
    [DisplayName("Clear Comments")]
    [Description("Clear the threaded comments.")]
    public InArgument<bool> ClearComments { get; set; } = null!;

    /// <summary>Clear the values and formulas.</summary>
    [Category("Input")]
    [DisplayName("Clear Contents")]
    [Description("Clear the values and formulas.")]
    public InArgument<bool> ClearContents { get; set; } = null!;

    /// <summary>Clear the formatting.</summary>
    [Category("Input")]
    [DisplayName("Clear Formats")]
    [Description("Clear the formatting.")]
    public InArgument<bool> ClearFormats { get; set; } = null!;

    /// <summary>Clear the hyperlinks.</summary>
    [Category("Input")]
    [DisplayName("Clear Hyperlinks")]
    [Description("Clear the hyperlinks.")]
    public InArgument<bool> ClearHyperlinks { get; set; } = null!;

    /// <summary>Clear the notes.</summary>
    [Category("Input")]
    [DisplayName("Clear Notes")]
    [Description("Clear the notes.")]
    public InArgument<bool> ClearNotes { get; set; } = null!;

    /// <summary>Clear the outline grouping.</summary>
    [Category("Input")]
    [DisplayName("Clear Outline")]
    [Description("Clear the outline grouping.")]
    public InArgument<bool> ClearOutline { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.ClearSheet(
            sheetName,
            CellRange?.Get(context) ?? string.Empty,
            new ClearOptions
            {
                All = ClearAll?.Get(context) ?? false,
                Contents = ClearContents?.Get(context) ?? false,
                Formats = ClearFormats?.Get(context) ?? false,
                Comments = ClearComments?.Get(context) ?? false,
                Notes = ClearNotes?.Get(context) ?? false,
                Hyperlinks = ClearHyperlinks?.Get(context) ?? false,
                Outline = ClearOutline?.Get(context) ?? false,
            });
}
