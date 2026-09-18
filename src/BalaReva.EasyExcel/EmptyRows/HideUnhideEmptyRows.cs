using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.EmptyRows;

/// <summary>Hides or shows the empty rows of a sheet.</summary>
[DisplayName("Hide Unhide Empty Rows")]
[Description("Hides or shows the empty rows of a sheet.")]
public sealed class HideUnhideEmptyRows : BaseActivity
{
    /// <summary>The rows this activity acted on, numbered from 1.</summary>
    [Category("Output")]
    [DisplayName("Affected Rows")]
    [Description("The rows this activity acted on, numbered from 1.")]
    public OutArgument<long[]> AffectedRows { get; set; } = null!;

    /// <summary>Whether to hide the rows or show them again.</summary>
    [Category("Input")]
    [DisplayName("Hide Delete")]
    [Description("Whether to hide the rows or show them again.")]
    public HideDeleteEnum HideDelete { get; set; } = HideDeleteEnum.Hide;

    /// <summary>Row to start looking from, numbered from 1.</summary>
    [Category("Input")]
    [DisplayName("Start Row Index")]
    [Description("Row to start looking from, numbered from 1.")]
    public InArgument<long> StartRowIndex { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => AffectedRows.Set(context, workbook.HideUnhideEmptyRows(
            sheetName, StartRowIndex?.Get(context) ?? 1, HideDelete));
}
