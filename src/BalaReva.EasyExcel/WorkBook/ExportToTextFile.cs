using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.WorkBook;

/// <summary>Writes a sheet out as a text file.</summary>
[DisplayName("Export To Text File")]
[Description("Writes a sheet out as a text file.")]
public sealed class ExportToTextFile : ExcelActivity
{
    /// <summary>Sheet to write out. Empty means the active sheet.</summary>
    [Category("Input")]
    [DisplayName("Sheet Name")]
    [Description("Sheet to write out. Empty means the active sheet.")]
    public InArgument<string> SheetName { get; set; } = null!;

    /// <summary>Full path of the text file to write.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Text File Path")]
    [Description("Full path of the text file to write.")]
    public InArgument<string> TextFilePath { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook)
        => workbook.ExportToTextFile(
            SheetName?.Get(context) ?? string.Empty,
            Require(context, TextFilePath, nameof(TextFilePath)));
}
