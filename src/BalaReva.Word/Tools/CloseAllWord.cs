using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word.Tools;

/// <summary>Closes every open Word document and quits Word.</summary>
/// <remarks>
/// Discards unsaved changes in every open document, not just the scope's own. The
/// published activity does the same; it exists to clean up after a run that left Word
/// processes behind.
/// </remarks>
[DisplayName("Close All Word")]
[Description("Closes every open Word document without saving, and quits Word.")]
public sealed class CloseAllWord : BaseNativeChild
{
    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.CloseAllWord();
}
