using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Deletes whole columns.</summary>
[DisplayName("Delete Columns")]
[Description("Deletes whole columns.")]
public sealed class DeleteColumns : BaseActivity
{
    /// <summary>Columns to delete, for example B or B:C.</summary>
    [Category("Input")]
    [DisplayName("Columns Range")]
    [Description("Columns to delete, for example B or B:C.")]
    public InArgument<string[]> ColumnsRange { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.DeleteColumns(sheetName, ColumnsRange?.Get(context) ?? []);
}
