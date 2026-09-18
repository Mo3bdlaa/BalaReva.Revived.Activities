using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Finds the last column of a sheet with anything in it.</summary>
[DisplayName("Find Last Column")]
[Description("Finds the last column of a sheet with anything in it.")]
public sealed class FindLastColumn : BaseActivity
{
    /// <summary>Its position, numbered from 1.</summary>
    [Category("Output")]
    [DisplayName("Column Index")]
    [Description("Its position, numbered from 1.")]
    public OutArgument<int> ColumnIndex { get; set; } = null!;

    /// <summary>Its letter.</summary>
    [Category("Output")]
    [DisplayName("Column Name")]
    [Description("Its letter.")]
    public OutArgument<string> ColumnName { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
    {
        var (index, name) = workbook.FindLastColumn(sheetName);
        ColumnIndex.Set(context, index);
        ColumnName.Set(context, name);
    }
}
