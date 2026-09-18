using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.Sheets;

/// <summary>Places an image anchored to a cell.</summary>
[DisplayName("Insert Image At Cell")]
[Description("Places an image anchored to a cell.")]
public sealed class InsertImageAtCell : ExcelCore
{
    /// <summary>Cell to anchor the image to.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell")]
    [Description("Cell to anchor the image to, for example B4.")]
    public InArgument<string> Cell { get; set; } = null!;

    /// <summary>Image to place.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Path")]
    [Description("Full path of the image to place.")]
    public InArgument<string> ImagePath { get; set; } = null!;

    /// <summary>Width in points.</summary>
    [Category("Input")]
    [DisplayName("Image Width")]
    [Description("Width in points. Zero uses the anchor cell's width.")]
    public InArgument<float> ImageWidth { get; set; } = null!;

    /// <summary>Height in points.</summary>
    [Category("Input")]
    [DisplayName("Image Height")]
    [Description("Height in points. Zero uses the anchor cell's height.")]
    public InArgument<float> ImageHeight { get; set; } = null!;

    /// <summary>True when the image was placed.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the image was placed.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.InsertImageAtCell(
            sheetName,
            Require(context, Cell, nameof(Cell)),
            Require(context, ImagePath, nameof(ImagePath)),
            ImageWidth?.Get(context) ?? 0,
            ImageHeight?.Get(context) ?? 0);

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
