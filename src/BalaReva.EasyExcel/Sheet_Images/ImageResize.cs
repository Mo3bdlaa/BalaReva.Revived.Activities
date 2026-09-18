using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheet_Images;

/// <summary>Resizes a picture on a sheet.</summary>
/// <remarks>
/// Both measurements are in points, and zero leaves that one alone.
/// </remarks>
[DisplayName("Image Resize")]
[Description("Resizes a picture on a sheet.")]
public sealed class ImageResize : BaseActivity
{
    /// <summary>Height in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Image Height")]
    [Description("Height in points. Zero leaves it alone.")]
    public InArgument<float> ImageHeight { get; set; } = null!;

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

    /// <summary>Width in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Image Width")]
    [Description("Width in points. Zero leaves it alone.")]
    public InArgument<float> ImageWidth { get; set; } = null!;

    /// <summary>True when the picture was found and resized.</summary>
    [Category("Output")]
    [DisplayName("Result")]
    [Description("True when the picture was found and resized.")]
    public OutArgument<bool> Result { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => Result.Set(context, workbook.ImageResize(
            new ImageRef
            {
                SheetName = sheetName,
                ImageName = ImageName?.Get(context) ?? string.Empty,
                ImageIndex = ImageIndex?.Get(context) ?? 0,
            },
            ImageWidth?.Get(context) ?? 0,
            ImageHeight?.Get(context) ?? 0));
}
