using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.AddIns;

/// <summary>Lists every Excel add-in by name and path.</summary>
[DisplayName("Get All Addins")]
[Description("Lists every Excel add-in by name and path.")]
public sealed class GetAllAddins : ExcelActivity
{
    /// <summary>Every add-in, keyed by name, valued by full path.</summary>
    [RequiredArgument]
    [Category("Output")]
    [DisplayName("Addins List")]
    [Description("Every add-in, keyed by name, valued by full path.")]
    public OutArgument<Dictionary<string, string>> AddinsList { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook)
        => AddinsList.Set(context, workbook.GetAllAddins());
}
