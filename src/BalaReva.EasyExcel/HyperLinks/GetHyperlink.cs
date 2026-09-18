using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.HyperLinks;

/// <summary>Reads the address a cell's hyperlink points at.</summary>
[DisplayName("Get Hyperlink")]
[Description("Reads the address a cell's hyperlink points at.")]
public sealed class GetHyperlink : BaseActivity
{
    /// <summary>Where the hyperlink points, or empty when the cell has none.</summary>
    [Category("Output")]
    [DisplayName("Address")]
    [Description("Where the hyperlink points, or empty when the cell has none.")]
    public OutArgument<string> Address { get; set; } = null!;

    /// <summary>Cell to read, for example B2.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell")]
    [Description("Cell to read, for example B2.")]
    public InArgument<string> Cell { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => Address.Set(context, workbook.GetHyperlink(
            sheetName, Require(context, Cell, nameof(Cell))));
}
