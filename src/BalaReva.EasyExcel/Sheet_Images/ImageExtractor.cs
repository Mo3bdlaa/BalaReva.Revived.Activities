using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.Sheet_Images;

/// <summary>Saves a sheet's pictures as image files.</summary>
[DisplayName("Image Extractor")]
[Description("Saves a sheet's pictures as image files.")]
public sealed class ImageExtractor : BaseActivity
{
    /// <summary>Image format to write.</summary>
    [Category("Input")]
    [DisplayName("File Extension")]
    [Description("Image format to write.")]
    public FileExtension FileExtension { get; set; } = FileExtension.Bmp;

    /// <summary>Folder to write the images into.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Folder")]
    [Description("Folder to write the images into.")]
    public InArgument<string> ImageFolder { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.ImageExtractor(
            sheetName,
            Require(context, ImageFolder, nameof(ImageFolder)),
            FileExtension);
}
