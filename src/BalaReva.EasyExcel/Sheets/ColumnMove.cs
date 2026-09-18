using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Moves a column range somewhere else on the sheet.</summary>
[DisplayName("Column Move")]
[Description("Moves a column range somewhere else on the sheet.")]
public sealed class ColumnMove : BaseActivity
{
    /// <summary>Where to move them, for example F:G.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Destination Col Range")]
    [Description("Where to move them, for example F:G.")]
    public InArgument<string> DestinationColRange { get; set; } = null!;

    /// <summary>Columns to move, for example B:C.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Source Col Range")]
    [Description("Columns to move, for example B:C.")]
    public InArgument<string> SourceColRange { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => workbook.ColumnMove(
            sheetName,
            Require(context, SourceColRange, nameof(SourceColRange)),
            Require(context, DestinationColRange, nameof(DestinationColRange)));
}
