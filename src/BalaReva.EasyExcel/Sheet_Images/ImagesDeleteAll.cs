using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheet_Images;

/// <summary>Deletes every picture on a sheet.</summary>
[DisplayName("Images Delete All")]
[Description("Deletes every picture on a sheet.")]
public sealed class ImagesDeleteAll : BaseActivity
{
    /// <summary>How many pictures were deleted.</summary>
    [Category("Output")]
    [DisplayName("Deleted Image Count")]
    [Description("How many pictures were deleted.")]
    public OutArgument<int> DeletedImageCount { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => DeletedImageCount.Set(context, workbook.ImagesDeleteAll(sheetName));
}
