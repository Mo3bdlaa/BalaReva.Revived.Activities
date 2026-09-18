using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Deletes whole rows.</summary>
[DisplayName("Delete Rows")]
[Description("Deletes whole rows.")]
public sealed class DeleteRows : BaseActivity
{
    /// <summary>Rows to delete, for example 3 or 3:7.</summary>
    [Category("Input")]
    [DisplayName("Row Range")]
    [Description("Rows to delete, for example 3 or 3:7.")]
    public InArgument<string[]> RowRange { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.DeleteRows(sheetName, RowRange?.Get(context) ?? []);
}
