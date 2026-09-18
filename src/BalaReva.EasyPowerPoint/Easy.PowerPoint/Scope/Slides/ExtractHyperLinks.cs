using System.Activities;
using System.ComponentModel;
using System.Data;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint;

namespace BalaReva.Easy.PowerPoint.Scope.Slides;

/// <summary>Lists the hyperlinks found on the given slides.</summary>
/// <remarks>
/// Note the namespace: the published package put this one and GetRowItem under
/// BalaReva.Easy.PowerPoint rather than BalaReva.EasyPowerPoint like everything
/// else. A workflow binds by full type name, so it stays where it is.
/// </remarks>
[DisplayName("Extract Hyper Links")]
[Description("Lists the hyperlinks found on the given slides.")]
public sealed class ExtractHyperLinks : BaseNativeChild
{
    /// <summary>Slides to search, numbered from 1. Empty means every slide.</summary>
    [Category("Input")]
    [DisplayName("Slide Indexes")]
    [Description("Slides to search, numbered from 1. Empty means every slide.")]
    public InArgument<int[]> SlideIndexes { get; set; } = null!;

    /// <summary>One row per hyperlink, with its slide, display text and address.</summary>
    [Category("Output")]
    [DisplayName("Result Table")]
    [Description("One row per hyperlink, with its slide, display text and address.")]
    public OutArgument<DataTable> ResultTable { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => ResultTable.Set(
            context, presentation.ExtractHyperLinks(SlideIndexes?.Get(context) ?? []));
}
