using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word.Pages;

/// <summary>Turns off displayed line numbers in every section.</summary>
[DisplayName("Remove Display Line Number")]
[Description("Turns off displayed line numbers in every section of the document.")]
public sealed class RemoveDisplayLineNumber : BaseNativeChild
{
    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.RemoveDisplayLineNumber();
}
