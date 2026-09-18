using System.Activities;
using System.ComponentModel;
using System.Data;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Finds text across the given slides.</summary>
[DisplayName("Find Text")]
[Description("Finds text across the given slides.")]
public sealed class FindText : BaseNativeChild
{
    /// <summary>Slides to search, numbered from 1. Empty means every slide.</summary>
    [Category("Input")]
    [DisplayName("Slide Indexes")]
    [Description("Slides to search, numbered from 1. Empty means every slide.")]
    public InArgument<int[]> SlideIndexes { get; set; } = null!;

    /// <summary>Text to find.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Find String")]
    [Description("Text to find.")]
    public InArgument<string> FindString { get; set; } = null!;

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

    /// <summary>The slides the text was found on.</summary>
    [Category("Output")]
    [DisplayName("Result Array")]
    [Description("The slides the text was found on, numbered from 1.")]
    public OutArgument<int[]> ResultArray { get; set; } = null!;

    /// <summary>One row per match, with its slide and shape.</summary>
    [Category("Output")]
    [DisplayName("Result Table")]
    [Description("One row per match, with its slide and shape.")]
    public OutArgument<DataTable> ResultTable { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
    {
        var (slides, table) = presentation.FindText(
            SlideIndexes?.Get(context) ?? [],
            Require(context, FindString, nameof(FindString)),
            MatchCase,
            WholeWord);

        ResultArray.Set(context, slides);
        ResultTable.Set(context, table);
    }
}
