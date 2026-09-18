using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.EmptyRows;

/// <summary>Finds the empty rows of a sheet.</summary>
[DisplayName("Find Empty Rows")]
[Description("Finds the empty rows of a sheet.")]
public sealed class FindEmptyRows : BaseActivity
{
    /// <summary>The rows this activity acted on, numbered from 1.</summary>
    [Category("Output")]
    [DisplayName("Affected Rows")]
    [Description("The rows this activity acted on, numbered from 1.")]
    public OutArgument<long[]> AffectedRows { get; set; } = null!;

    /// <summary>Row to start looking from, numbered from 1.</summary>
    [Category("Input")]
    [DisplayName("Start Row Index")]
    [Description("Row to start looking from, numbered from 1.")]
    public InArgument<long> StartRowIndex { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => AffectedRows.Set(
            context, workbook.FindEmptyRows(sheetName, StartRowIndex?.Get(context) ?? 0));
}
