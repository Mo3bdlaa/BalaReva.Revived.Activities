using System.Activities;
using System.ComponentModel;
using System.Drawing;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.WorkBook;

/// <summary>Colours a sheet's tab.</summary>
[DisplayName("Tab Color")]
[Description("Colours a sheet's tab.")]
public sealed class TabColor : BaseActivity
{
    /// <summary>Colour for the tab.</summary>
    [Category("Input")]
    [DisplayName("Tab Color")]
    [Description("Colour for the tab.")]
    public InArgument<Color> Tab_Color { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.TabColor(sheetName, Tab_Color?.Get(context) ?? Color.Empty);
}
