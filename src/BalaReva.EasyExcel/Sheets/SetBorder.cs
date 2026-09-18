using System.Activities;
using System.ComponentModel;
using System.Drawing;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;
using Interop = Microsoft.Office.Interop.Excel;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Draws borders around or through a range.</summary>
/// <remarks>
/// LineStyle is an XlLineStyle, straight from the Excel interop assembly. That is how the
/// published package declared it, and a workflow binds by type, so it stays: this is the
/// one activity in the family that puts an interop type on an activity.
/// </remarks>
[DisplayName("Set Border")]
[Description("Draws borders around or through a range.")]
public sealed class SetBorder : BaseActivity
{
    /// <summary>Line colour. Empty leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Border Color")]
    [Description("Line colour. Empty leaves it alone.")]
    public InArgument<Color> BorderColor { get; set; } = null!;

    /// <summary>Line thickness. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Border Weight")]
    [Description("Line thickness. Zero leaves it alone.")]
    public InArgument<double> BorderWeight { get; set; } = null!;

    /// <summary>Range to work over, for example A1:D10.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell Range")]
    [Description("Range to work over, for example A1:D10.")]
    public InArgument<string> CellRange { get; set; } = null!;

    /// <summary>Line style to draw with.</summary>
    [Category("Input")]
    [DisplayName("Line Style")]
    [Description("Line style to draw with.")]
    public Interop.XlLineStyle LineStyle { get; set; }

    /// <summary>Which edges to draw.</summary>
    [Category("Input")]
    [DisplayName("Presets")]
    [Description("Which edges to draw.")]
    public BorderEnum Presets { get; set; } = BorderEnum.NoBorder;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.SetBorder(
            sheetName,
            Require(context, CellRange, nameof(CellRange)),
            Presets,
            LineStyle,
            BorderWeight?.Get(context) ?? 0,
            BorderColor?.Get(context) ?? Color.Empty);
}
