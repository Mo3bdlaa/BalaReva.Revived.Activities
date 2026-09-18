using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.TableArea;

/// <summary>Deletes a row from a table.</summary>
[DisplayName("Delete Row")]
[Description("Deletes a row from a table.")]
public sealed class DeleteRow : BaseTableNativeChild
{
    /// <summary>Which row to delete. Rows are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Row Index")]
    [Description("Which row to delete. Rows are numbered from 1.")]
    public InArgument<int> RowIndex { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.DeleteRow(Table(context), RowIndex.Get(context));
}
