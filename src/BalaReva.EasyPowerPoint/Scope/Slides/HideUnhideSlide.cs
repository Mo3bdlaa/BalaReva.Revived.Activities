using System.Activities;
using System.ComponentModel;
using BalaReva.Easy.PowerPoint.Utilities;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Shows or hides a slide during a slide show.</summary>
[DisplayName("Hide Unhide Slide")]
[Description("Shows or hides a slide during a slide show.")]
public sealed class HideUnhideSlide : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>Whether the slide is hidden during a slide show.</summary>
    [Category("Input")]
    [DisplayName("Slide Show")]
    [Description("Whether the slide is hidden during a slide show.")]
    public HideUnhideEnum SlideShow { get; set; } = HideUnhideEnum.Hide;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.HideUnhideSlide(SlideIndex.Get(context), SlideShow);
}
