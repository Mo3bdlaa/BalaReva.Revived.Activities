using System.Activities;

namespace BalaReva.EasyExcel;

/// <summary>
/// The execution property an <c>ExcelScope</c> publishes so its children can find the
/// open workbook.
/// </summary>
/// <remarks>
/// Derives from <see cref="Handle"/> because that is the constraint on
/// <see cref="CodeActivityContext.GetProperty{THandle}"/>, which is how an
/// <c>ExcelActivity</c> resolves its scope.
/// </remarks>
public sealed class ExcelScopeHandle : Handle
{
    /// <summary>The workbook the enclosing scope has open.</summary>
    public IExcelWorkbook? Workbook { get; set; }
}
