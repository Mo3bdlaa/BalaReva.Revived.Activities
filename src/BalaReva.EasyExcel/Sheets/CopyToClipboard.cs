using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Copies a range to the clipboard.</summary>
[DisplayName("Copy To Clipboard")]
[Description("Copies a range to the clipboard.")]
public sealed class CopyToClipboard : BaseActivity
{
    /// <summary>Range to copy, for example A1:D10.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell")]
    [Description("Range to copy, for example A1:D10.")]
    public InArgument<string> Cell { get; set; } = null!;

    /// <summary>Copy only the rows a filter is showing.</summary>
    [Category("Input")]
    [DisplayName("Read Filter")]
    [Description("Copy only the rows a filter is showing.")]
    public bool ReadFilter { get; set; }

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.CopyToClipboard(
            sheetName, Require(context, Cell, nameof(Cell)), ReadFilter);
}
