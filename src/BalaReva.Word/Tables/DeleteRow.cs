using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word.Tables;

/// <summary>Deletes a row from a table.</summary>
[DisplayName("Delete Row")]
[Description("Deletes a row from a table.")]
public sealed class DeleteRow : BaseNativeChild
{
    /// <summary>Which table. Tables are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Table Index")]
    [Description("Which table. Tables are numbered from 1.")]
    public InArgument<int> TableIndex { get; set; } = null!;

    /// <summary>Which row. Rows are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Row Index")]
    [Description("Which row to delete. Rows are numbered from 1.")]
    public InArgument<int> RowIndex { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.DeleteRow(TableIndex.Get(context), RowIndex.Get(context));
}
