using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Charts;

/// <summary>Copies a chart to the clipboard.</summary>
[DisplayName("Chart Copy To Clipboard")]
[Description("Copies a chart to the clipboard.")]
public sealed class ChartCopyToClipboard : BaseActivity
{
    /// <summary>Position of the chart on the sheet, numbered from 1.</summary>
    [Category("Input")]
    [DisplayName("Chart Index")]
    [Description("Position of the chart on the sheet, numbered from 1.")]
    public InArgument<int> ChartIndex { get; set; } = null!;

    /// <summary>Name of the chart. Takes precedence over the index.</summary>
    [Category("Input")]
    [DisplayName("Chart Name")]
    [Description("Name of the chart. Takes precedence over the index.")]
    public InArgument<string> ChartName { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.ChartCopyToClipboard(
            new ChartRef
            {
                SheetName = sheetName,
                ChartName = ChartName?.Get(context) ?? string.Empty,
                ChartIndex = ChartIndex?.Get(context) ?? 0,
            });
}
