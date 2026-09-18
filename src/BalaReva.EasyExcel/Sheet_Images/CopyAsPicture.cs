using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheet_Images;

/// <summary>Copies a range to an image file.</summary>
[DisplayName("Copy As Picture")]
[Description("Copies a range to an image file.")]
public sealed class CopyAsPicture : BaseActivity
{
    /// <summary>Range to copy, for example A1:D10.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell")]
    [Description("Range to copy, for example A1:D10.")]
    public InArgument<string> Cell { get; set; } = null!;

    /// <summary>Full path of the image file to write.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image File Path")]
    [Description("Full path of the image file to write.")]
    public InArgument<string> ImageFilePath { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.CopyAsPicture(
            sheetName,
            Require(context, Cell, nameof(Cell)),
            Require(context, ImageFilePath, nameof(ImageFilePath)));
}
