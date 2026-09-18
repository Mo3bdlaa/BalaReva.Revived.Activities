using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.Charts;

/// <summary>Saves a sheet's charts as image files.</summary>
[DisplayName("Chart Image Extract")]
[Description("Saves a sheet's charts as image files.")]
public sealed class ChartImageExtract : BaseActivity
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
        => workbook.ChartImageExtract(
            sheetName,
            Require(context, ImageFolder, nameof(ImageFolder)),
            FileExtension);
}
