using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.WorkBook;

/// <summary>Lists the sheets that are visible, or the ones that are not.</summary>
[DisplayName("Show Visible Sheet")]
[Description("Lists the sheets that are visible, or the ones that are not.")]
public sealed class ShowVisibleSheet : ExcelActivity
{
    /// <summary>The sheets, by name.</summary>
    [Category("Output")]
    [DisplayName("Sheet List")]
    [Description("The sheets, by name.")]
    public OutArgument<string[]> SheetList { get; set; } = null!;

    /// <summary>Whether to list the visible sheets or the hidden ones.</summary>
    [Category("Input")]
    [DisplayName("Visible Type")]
    [Description("Whether to list the visible sheets or the hidden ones.")]
    public VisibleInvisibleEnum VisibleType { get; set; } = VisibleInvisibleEnum.Visible;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook)
        => SheetList.Set(context, workbook.ShowVisibleSheet(VisibleType));
}
