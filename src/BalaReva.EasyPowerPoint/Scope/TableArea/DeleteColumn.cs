using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.TableArea;

/// <summary>Deletes a column from a table.</summary>
[DisplayName("Delete Column")]
[Description("Deletes a column from a table.")]
public sealed class DeleteColumn : BaseTableNativeChild
{
    /// <summary>Which column to delete. Columns are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Column Index")]
    [Description("Which column to delete. Columns are numbered from 1.")]
    public InArgument<int> ColumnIndex { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.DeleteColumn(Table(context), ColumnIndex.Get(context));
}
