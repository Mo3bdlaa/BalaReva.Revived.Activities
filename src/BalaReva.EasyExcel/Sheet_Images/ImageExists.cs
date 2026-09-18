using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheet_Images;

/// <summary>Reports whether a picture is on a sheet.</summary>
[DisplayName("Image Exists")]
[Description("Reports whether a picture is on a sheet.")]
public sealed class ImageExists : BaseActivity
{
    /// <summary>Position of the picture on the sheet, numbered from 1.</summary>
    [Category("Input")]
    [DisplayName("Image Index")]
    [Description("Position of the picture on the sheet, numbered from 1.")]
    public InArgument<int> ImageIndex { get; set; } = null!;

    /// <summary>Name of the picture. Takes precedence over the index.</summary>
    [Category("Input")]
    [DisplayName("Image Name")]
    [Description("Name of the picture. Takes precedence over the index.")]
    public InArgument<string> ImageName { get; set; } = null!;

    /// <summary>True when the picture is there.</summary>
    [Category("Output")]
    [DisplayName("Result")]
    [Description("True when the picture is there.")]
    public OutArgument<bool> Result { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => Result.Set(context, workbook.ImageExists(
            new ImageRef
            {
                SheetName = sheetName,
                ImageName = ImageName?.Get(context) ?? string.Empty,
                ImageIndex = ImageIndex?.Get(context) ?? 0,
            }));
}
