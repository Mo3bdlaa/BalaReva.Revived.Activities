using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word.Tables;

/// <summary>Deletes a whole table.</summary>
[DisplayName("Delete Table")]
[Description("Deletes a whole table from the document.")]
public sealed class DeleteTable : BaseNativeChild
{
    /// <summary>Which table. Tables are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Table Index")]
    [Description("Which table to delete. Tables are numbered from 1.")]
    public InArgument<int> TableIndex { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.DeleteTable(TableIndex.Get(context));
}
