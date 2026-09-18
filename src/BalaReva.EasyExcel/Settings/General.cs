using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Settings;

/// <summary>Turns Excel's adaptive menus on or off.</summary>
[DisplayName("General")]
[Description("Turns Excel's adaptive menus on or off.")]
public sealed class General : ExcelActivity
{
    /// <summary>Whether Excel shortens its menus to recently used commands.</summary>
    [Category("Input")]
    [DisplayName("Adaptive Menus")]
    [Description("Whether Excel shortens its menus to recently used commands.")]
    public bool Adaptive_Menus { get; set; }

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook)
        => workbook.General(Adaptive_Menus);
}
