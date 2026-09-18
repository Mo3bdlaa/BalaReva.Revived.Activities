using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint;

namespace BalaReva.Easy.PowerPoint.Scope.TableArea;

/// <summary>Reads the text of one table cell.</summary>
/// <remarks>
/// In the BalaReva.Easy.PowerPoint namespace rather than BalaReva.EasyPowerPoint,
/// as the published package had it.
/// </remarks>
[DisplayName("Get Row Item")]
[Description("Reads the text of one table cell.")]
public sealed class GetRowItem : BaseTableNativeChild
{
    /// <summary>Which row. Rows are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Row Index")]
    [Description("Which row. Rows are numbered from 1.")]
    public InArgument<int> RowIndex { get; set; } = null!;

    /// <summary>Which column. Columns are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Column Index")]
    [Description("Which column. Columns are numbered from 1.")]
    public InArgument<int> ColumnIndex { get; set; } = null!;

    /// <summary>The text in the cell.</summary>
    [Category("Output")]
    [DisplayName("Value")]
    [Description("The text in the cell.")]
    public OutArgument<string> Value { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => Value.Set(context, presentation.GetRowItem(
            Table(context), RowIndex.Get(context), ColumnIndex.Get(context)));
}
