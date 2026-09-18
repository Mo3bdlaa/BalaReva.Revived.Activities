using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Counts the slides in the presentation.</summary>
[DisplayName("Slide Count")]
[Description("Counts the slides in the presentation.")]
public sealed class SlideCount : BaseNativeChild
{
    /// <summary>Number of slides.</summary>
    [Category("Output")]
    [DisplayName("Total Slides")]
    [Description("Number of slides.")]
    public OutArgument<int> TotalSlides { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => TotalSlides.Set(context, presentation.SlideCount());
}
