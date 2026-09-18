using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.HyperLinks;

/// <summary>Puts a hyperlink on a cell.</summary>
[DisplayName("Insert Hyperlink")]
[Description("Puts a hyperlink on a cell.")]
public sealed class InsertHyperlink : BaseActivity
{
    /// <summary>Where the link should point.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Address")]
    [Description("Where the link should point.")]
    public InArgument<string> Address { get; set; } = null!;

    /// <summary>Cell to put the link on, for example B2.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell")]
    [Description("Cell to put the link on, for example B2.")]
    public InArgument<string> Cell { get; set; } = null!;

    /// <summary>Text to show in the cell. Empty keeps the cell's own text.</summary>
    [Category("Input")]
    [DisplayName("Display Text")]
    [Description("Text to show in the cell. Empty keeps the cell's own text.")]
    public InArgument<string> DisplayText { get; set; } = null!;

    /// <summary>Replace the text already in the cell.</summary>
    [Category("Input")]
    [DisplayName("Display Text Overwirte")]
    [Description("Replace the text already in the cell.")]
    public InArgument<bool> DisplayTextOverwirte { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.InsertHyperlink(
            sheetName,
            Require(context, Cell, nameof(Cell)),
            Require(context, Address, nameof(Address)),
            DisplayText?.Get(context) ?? string.Empty,
            DisplayTextOverwirte?.Get(context) ?? false);
}
