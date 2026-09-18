using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.TableArea;

/// <summary>Writes text into one table cell.</summary>
[DisplayName("Edit Table")]
[Description("Writes text into one table cell.")]
public sealed class EditTable : BaseTableNativeChild
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

    /// <summary>Text to write into the cell.</summary>
    [Category("Input")]
    [DisplayName("Cell Value")]
    [Description("Text to write into the cell.")]
    public InArgument<string> CellValue { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.EditTable(
            Table(context),
            RowIndex.Get(context),
            ColumnIndex.Get(context),
            CellValue?.Get(context) ?? string.Empty);
}
