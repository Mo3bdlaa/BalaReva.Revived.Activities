using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word.Documents;

/// <summary>Creates an empty document at the scope's path.</summary>
[DisplayName("Create Document")]
[Description("Creates an empty Word document at the path the scope was given.")]
public sealed class CreateDocument : BaseNativeChild
{
    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.CreateDocument();
}
