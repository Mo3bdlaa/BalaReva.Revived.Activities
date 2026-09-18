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
    public InArgument<bool> MatchCase { get; set; } = null!;

    /// <summary>Match whole words only.</summary>
    [Category("Input")]
    [DisplayName("Whole Word")]
    [Description("Match whole words only.")]
    public InArgument<bool> WholeWord { get; set; } = null!;

    /// <summary>The matching text, one entry each.</summary>
    [Category("Output")]
    [DisplayName("Result Array")]
    [Description("The matching text, one entry each.")]
    public OutArgument<string[]> ResultArray { get; set; } = null!;

    /// <summary>One row per match, with its slide and shape.</summary>
    [Category("Output")]
    [DisplayName("Result Table")]
    [Description("One row per match, with its slide and shape.")]
    public OutArgument<DataTable> ResultTable { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
    {
        var (array, table) = presentation.FindText(
            SlideIndexes?.Get(context) ?? [],
            Require(context, FindString, nameof(FindString)),
            MatchCase?.Get(context) ?? false,
            WholeWord?.Get(context) ?? false);

        ResultArray.Set(context, array);
        ResultTable.Set(context, table);
    }
}
