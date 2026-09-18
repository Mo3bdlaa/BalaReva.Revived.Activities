using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word.Documents;

/// <summary>Appends other documents onto a Word file.</summary>
/// <remarks>
/// Works on a file directly rather than inside a scope, so it carries its own file and
/// password arguments. Each source starts on a new page.
/// </remarks>
[DisplayName("Merge Documents")]
[Description("Appends other Word documents onto a file, each starting on a new page.")]
public sealed class MergeDocuments : BaseWord
{
    /// <summary>Documents to append, in order.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Source Files")]
    [Description("Full paths of the documents to append, in order.")]
    public InArgument<string[]> SourceFiles { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordService service)
    {
        var sources = SourceFiles?.Get(context);
        if (sources is null || sources.Length == 0)
            throw new ArgumentException("SourceFiles is required.", nameof(SourceFiles));

        service.MergeDocuments(
            Require(context, WordFile, nameof(WordFile)),
            OpenPassword?.Get(context) ?? string.Empty,
            ModifyPassword?.Get(context) ?? string.Empty,
            sources);
    }
}
