using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Finds the last row of a sheet with anything in it.</summary>
[DisplayName("Find Last Row")]
[Description("Finds the last row of a sheet with anything in it.")]
public sealed class FindLastRow : BaseActivity
{
    /// <summary>Its number, counted from 1.</summary>
    [Category("Output")]
    [DisplayName("Row Index")]
    [Description("Its number, counted from 1.")]
    public OutArgument<int> RowIndex { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => RowIndex.Set(context, workbook.FindLastRow(sheetName));
}
