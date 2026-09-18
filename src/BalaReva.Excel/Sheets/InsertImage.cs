using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;
using BalaReva.Excel.Utilities;

namespace BalaReva.Excel.Sheets;

/// <summary>Places an image on a sheet.</summary>
[DisplayName("Insert Image")]
[Description("Places an image on a sheet at a given position and size.")]
public sealed class InsertImage : ExcelCore
{
    /// <summary>Image to place.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Path")]
    [Description("Full path of the image to place.")]
    public InArgument<string> ImagePath { get; set; } = null!;

    /// <summary>Where and how big.</summary>
    [Category("Input")]
    [DisplayName("Size")]
    [Description("Where the image sits on the sheet and how big it is, in points.")]
    public ObjectSize Size { get; set; } = new();

    /// <summary>True when the image was placed.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the image was placed.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.InsertImage(
            sheetName, Require(context, ImagePath, nameof(ImagePath)), Size ?? new ObjectSize());

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
