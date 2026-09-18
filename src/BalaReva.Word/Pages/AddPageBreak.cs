using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word.Pages;

/// <summary>Inserts a page break at the current selection.</summary>
[DisplayName("Add Page Break")]
[Description("Inserts a page break at the current selection.")]
public sealed class AddPageBreak : BaseNativeChild
{
    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.AddPageBreak();
}
