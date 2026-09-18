using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word.Tables;

/// <summary>Copies a table to the clipboard.</summary>
[DisplayName("Copy Table To Clipboard")]
[Description("Copies a table to the Windows clipboard.")]
public sealed class CopyTableToClipboard : BaseNativeChild
{
    /// <summary>Which table. Tables are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Table Index")]
    [Description("Which table to copy. Tables are numbered from 1.")]
    public InArgument<int> TableIndex { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.CopyTableToClipboard(TableIndex.Get(context));
}
