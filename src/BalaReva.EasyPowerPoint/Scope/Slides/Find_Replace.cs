using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Replaces text across the given slides.</summary>
/// <remarks>
/// The class name keeps its underscore, as the published package had it.
/// </remarks>
[DisplayName("Find Replace")]
[Description("Replaces text across the given slides.")]
public sealed class Find_Replace : BaseNativeChild
{
    /// <summary>Slides to work on, numbered from 1. Empty means every slide.</summary>
    [Category("Input")]
    [DisplayName("Slide Indexes")]
    [Description("Slides to work on, numbered from 1. Empty means every slide.")]
    public InArgument<int[]> SlideIndexes { get; set; } = null!;

    /// <summary>Text to find.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Find Text")]
    [Description("Text to find.")]
    public InArgument<string> FindText { get; set; } = null!;

    /// <summary>Text to put in its place. Empty deletes the match.</summary>
    [Category("Input")]
    [DisplayName("Replace Text")]
    [Description("Text to put in its place. Empty deletes the match.")]
    public InArgument<string> ReplaceText { get; set; } = null!;

    /// <summary>Whether the search is case sensitive.</summary>
    [Category("Input")]
    [DisplayName("Match Case")]
    [Description("Whether the search is case sensitive.")]
    public bool MatchCase { get; set; }

    /// <summary>Match whole words only.</summary>
    [Category("Input")]
    [DisplayName("Whole Word")]
    [Description("Match whole words only.")]
    public bool WholeWord { get; set; }

    /// <summary>Replace only the first match in each shape.</summary>
    [Category("Input")]
    [DisplayName("First Occurrence")]
    [Description("Replace only the first match in each shape.")]
    public bool FirstOccurrence { get; set; }

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.FindReplace(
            SlideIndexes?.Get(context) ?? [],
            Require(context, FindText, nameof(FindText)),
            ReplaceText?.Get(context) ?? string.Empty,
            MatchCase,
            WholeWord,
            FirstOccurrence);
}
