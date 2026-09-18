using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word.Tables;

/// <summary>Deletes a column from a table.</summary>
[DisplayName("Delete Column")]
[Description("Deletes a column from a table.")]
public sealed class DeleteColumn : BaseNativeChild
{
    /// <summary>Which table. Tables are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Table Index")]
    [Description("Which table. Tables are numbered from 1.")]
    public InArgument<int> TableIndex { get; set; } = null!;

    /// <summary>Which column. Columns are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Column Index")]
    [Description("Which column to delete. Columns are numbered from 1.")]
    public InArgument<int> ColumnIndex { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.DeleteColumn(TableIndex.Get(context), ColumnIndex.Get(context));
}
