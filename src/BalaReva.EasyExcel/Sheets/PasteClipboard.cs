using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Pastes the clipboard at a cell.</summary>
[DisplayName("Paste Clipboard")]
[Description("Pastes the clipboard at a cell.")]
public sealed class PasteClipboard : BaseActivity
{
    /// <summary>Cell to paste at, for example B2.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell")]
    [Description("Cell to paste at, for example B2.")]
    public InArgument<string> Cell { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.PasteClipboard(
            sheetName, Require(context, Cell, nameof(Cell)));
}
