using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Counts the images on a slide.</summary>
/// <remarks>
/// The output is called ChartCount, not ImageCount. That is how the published
/// package named it, and a workflow binds by property name.
/// </remarks>
[DisplayName("Image Shape Count")]
[Description("Counts the images on a slide.")]
public sealed class ImageShapeCount : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>Number of images on the slide.</summary>
    [Category("Output")]
    [DisplayName("Chart Count")]
    [Description("Number of images on the slide.")]
    public OutArgument<int> ChartCount { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => ChartCount.Set(context, presentation.ImageShapeCount(SlideIndex.Get(context)));
}
