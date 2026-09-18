using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word.Pages;

/// <summary>Reads Word's own counts for the document.</summary>
[DisplayName("Word Statistics")]
[Description("Reads Word's counts of pages, words, lines, paragraphs and characters.")]
public sealed class WordStatistics : BaseNativeChild
{
    /// <summary>Pages.</summary>
    [Category("Output")]
    [DisplayName("Pages")]
    [Description("Number of pages.")]
    public OutArgument<long> Pages { get; set; } = null!;

    /// <summary>Words.</summary>
    [Category("Output")]
    [DisplayName("Word Count")]
    [Description("Number of words.")]
    public OutArgument<long> WordCount { get; set; } = null!;

    /// <summary>Lines.</summary>
    [Category("Output")]
    [DisplayName("Lines")]
    [Description("Number of lines.")]
    public OutArgument<long> Lines { get; set; } = null!;

    /// <summary>Paragraphs.</summary>
    [Category("Output")]
    [DisplayName("Paragraphs")]
    [Description("Number of paragraphs.")]
    public OutArgument<long> Paragraphs { get; set; } = null!;

    /// <summary>Characters, not counting spaces.</summary>
    [Category("Output")]
    [DisplayName("Characters")]
    [Description("Number of characters, not counting spaces.")]
    public OutArgument<long> Characters { get; set; } = null!;

    /// <summary>Characters, counting spaces.</summary>
    [Category("Output")]
    [DisplayName("Characters With Spaces")]
    [Description("Number of characters, counting spaces.")]
    public OutArgument<long> CharactersWithSpaces { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document)
    {
        var statistics = document.Statistics();
        Pages.Set(context, statistics.Pages);
        WordCount.Set(context, statistics.WordCount);
        Lines.Set(context, statistics.Lines);
        Paragraphs.Set(context, statistics.Paragraphs);
        Characters.Set(context, statistics.Characters);
        CharactersWithSpaces.Set(context, statistics.CharactersWithSpaces);
    }
}
