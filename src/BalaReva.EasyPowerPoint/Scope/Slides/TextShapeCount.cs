using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Counts the text shapes on a slide.</summary>
/// <remarks>
/// The output is called ImageCount, not TextCount - the published package has
/// this one and ImageShapeCount named the other way round. Both are kept as
/// shipped, because a workflow binds by property name.
/// </remarks>
[DisplayName("Text Shape Count")]
[Description("Counts the text shapes on a slide.")]
public sealed class TextShapeCount : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>Number of text shapes on the slide.</summary>
    [Category("Output")]
    [DisplayName("Image Count")]
    [Description("Number of text shapes on the slide.")]
    public OutArgument<int> ImageCount { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => ImageCount.Set(context, presentation.TextShapeCount(SlideIndex.Get(context)));
}
