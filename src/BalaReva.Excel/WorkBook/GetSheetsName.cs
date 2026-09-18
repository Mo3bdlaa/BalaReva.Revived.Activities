using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.WorkBook;

/// <summary>Lists the workbook's sheet names.</summary>
/// <remarks>
/// Declares its own <c>SheetName</c> argument, shadowing the base's. That is how the
/// published package had it, and the name has to stay for workflows to keep binding.
/// </remarks>
[DisplayName("Get Sheets Name")]
[Description("Lists the names of every sheet in the workbook.")]
public sealed class GetSheetsName : ExcelCore
{
    /// <summary>Names of every sheet.</summary>
    [Category("Output")]
    [DisplayName("Sheets Name")]
    [Description("Names of every sheet in the workbook, in order.")]
    public OutArgument<string[]> SheetsName { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        SheetsName.Set(context, workbook.GetSheetNames());
}
