using System.Activities;
using System.ComponentModel;
using System.Data;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Settings.TrustedLocation;

/// <summary>Lists Excel's trusted locations.</summary>
[DisplayName("List Trusted Location")]
[Description("Lists Excel's trusted locations.")]
public sealed class ListTrustedLocation : ExcelActivity
{
    /// <summary>The trusted folders, by path.</summary>
    [Category("Output")]
    [DisplayName("Output List")]
    [Description("The trusted folders, by path.")]
    public OutArgument<string[]> OutputList { get; set; } = null!;

    /// <summary>One row per trusted folder, with its path and description.</summary>
    [Category("Output")]
    [DisplayName("Url Table")]
    [Description("One row per trusted folder, with its path and description.")]
    public OutArgument<DataTable> UrlTable { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook)
    {
        var (paths, table) = workbook.ListTrustedLocation();
        OutputList.Set(context, paths);
        UrlTable.Set(context, table);
    }
}
