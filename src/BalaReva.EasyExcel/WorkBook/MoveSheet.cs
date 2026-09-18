using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.WorkBook;

/// <summary>Moves a sheet in front of another.</summary>
[DisplayName("Move Sheet")]
[Description("Moves a sheet in front of another.")]
public sealed class MoveSheet : ExcelActivity
{
    /// <summary>Sheet to move it in front of. Empty moves it to the end.</summary>
    [Category("Input")]
    [DisplayName("Before Sheet")]
    [Description("Sheet to move it in front of. Empty moves it to the end.")]
    public InArgument<string> BeforeSheet { get; set; } = null!;

    /// <summary>Sheet to move.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Sheet Name")]
    [Description("Sheet to move.")]
    public InArgument<string> SheetName { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook)
        => workbook.MoveSheet(
            Require(context, SheetName, nameof(SheetName)),
            BeforeSheet?.Get(context) ?? string.Empty);
}
