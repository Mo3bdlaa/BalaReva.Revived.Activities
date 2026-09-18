using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;
using BalaReva.Excel.Utilities;

namespace BalaReva.Excel.Sheets;

/// <summary>Saves each chart on a sheet as an image.</summary>
[DisplayName("Extract Graph Image")]
[Description("Saves each chart on a sheet as an image file.")]
public sealed class ExtractGraphImage : ExcelCore
{
    /// <summary>Folder to save the images into.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Folder")]
    [Description("Folder to save the images into. Created if it does not exist.")]
    public InArgument<string> ImageFolder { get; set; } = null!;

    /// <summary>Image format to save in.</summary>
    [Category("Input")]
    [DisplayName("File Extension")]
    [Description("Format to save the images in.")]
    public FileExtensEnum FileExtension { get; set; } = FileExtensEnum.Png;

    /// <summary>True when the charts were saved.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the charts were saved.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.ExtractGraphImage(
            sheetName, Require(context, ImageFolder, nameof(ImageFolder)), FileExtension);

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
