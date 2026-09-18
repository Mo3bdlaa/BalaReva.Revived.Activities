using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.AddIns;

/// <summary>Reports whether an Excel add-in is installed.</summary>
[DisplayName("Exists AddIns")]
[Description("Reports whether an Excel add-in is installed.")]
public sealed class ExistsAddIns : ExcelActivity
{
    /// <summary>Name of the add-in to look for.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Add Ins Name")]
    [Description("Name of the add-in to look for.")]
    public InArgument<string> AddInsName { get; set; } = null!;

    /// <summary>True when an add-in of that name is installed.</summary>
    [RequiredArgument]
    [Category("Output")]
    [DisplayName("Result")]
    [Description("True when an add-in of that name is installed.")]
    public OutArgument<bool> Result { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook)
        => Result.Set(context, workbook.ExistsAddIns(
            Require(context, AddInsName, nameof(AddInsName))));
}
