using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheet_Images;

/// <summary>Deletes the named or numbered pictures on a sheet.</summary>
[DisplayName("Images Delete")]
[Description("Deletes the named or numbered pictures on a sheet.")]
public sealed class ImagesDelete : BaseActivity
{
    /// <summary>How many pictures were deleted.</summary>
    [Category("Output")]
    [DisplayName("Deleted Image Count")]
    [Description("How many pictures were deleted.")]
    public OutArgument<int> DeletedImageCount { get; set; } = null!;

    /// <summary>Pictures to delete, by position, numbered from 1.</summary>
    [Category("Input")]
    [DisplayName("Image Indexes")]
    [Description("Pictures to delete, by position, numbered from 1.")]
    public InArgument<int[]> ImageIndexes { get; set; } = null!;

    /// <summary>Pictures to delete, by name.</summary>
    [Category("Input")]
    [DisplayName("Image Names")]
    [Description("Pictures to delete, by name.")]
    public InArgument<string[]> ImageNames { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => DeletedImageCount.Set(context, workbook.ImagesDelete(
            sheetName,
            ImageNames?.Get(context) ?? [],
            ImageIndexes?.Get(context) ?? []));
}
