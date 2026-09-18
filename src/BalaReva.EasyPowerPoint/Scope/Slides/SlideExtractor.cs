using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.PowerPoint;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Reads a slide's text shapes: their text, font and position.</summary>
[DisplayName("Slide Extractor")]
[Description("Reads a slide's text shapes: their text, font and position.")]
public sealed class SlideExtractor : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>The slide's text shapes.</summary>
    [Category("Output")]
    [DisplayName("Slide Result")]
    [Description("The slide's text shapes.")]
    public OutArgument<SlideObject> SlideResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => SlideResult.Set(context, presentation.SlideExtractor(SlideIndex.Get(context)));
}
