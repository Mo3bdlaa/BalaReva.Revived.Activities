using System.Activities;
using System.ComponentModel;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.Pages;

/// <summary>Finds and replaces text in the document.</summary>
[DisplayName("Find Replace")]
[Description("Finds text in the document and replaces it.")]
public sealed class FindReplace : BaseNativeChild
{
    /// <summary>Text to find.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Find Text")]
    [Description("Text to find.")]
    public InArgument<string> FindText { get; set; } = null!;

    /// <summary>Text to put in its place.</summary>
    [Category("Input")]
    [DisplayName("Replace Text")]
    [Description("Text to put in its place. Empty deletes the match.")]
    public InArgument<string> ReplaceText { get; set; } = null!;

    /// <summary>Whether to match whole words only.</summary>
    [Category("Input")]
    [DisplayName("Find Option")]
    [Description("Whole matches complete words only; Part matches inside words too.")]
    public EnumFindReplace FindOption { get; set; } = EnumFindReplace.Part;

    /// <summary>Whether to replace one match or all of them.</summary>
    [Category("Input")]
    [DisplayName("Replace Option")]
    [Description("Replace the first match only, or every match.")]
    public EnumReplaceOption ReplaceOption { get; set; } = EnumReplaceOption.ReplaceAll;

    /// <summary>Whether the search is case sensitive.</summary>
    [Category("Input")]
    [DisplayName("Match Case")]
    [Description("Whether the search is case sensitive.")]
    public bool MatchCase { get; set; }

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.FindReplace(
            Require(context, FindText, nameof(FindText)),
            ReplaceText?.Get(context) ?? string.Empty,
            FindOption,
            ReplaceOption,
            MatchCase);
}
