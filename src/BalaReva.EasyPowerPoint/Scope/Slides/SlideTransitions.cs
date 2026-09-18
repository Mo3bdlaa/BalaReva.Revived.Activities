using System.Activities;
using System.ComponentModel;
using BalaReva.Easy.PowerPoint.Utilities;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Sets a slide transition.</summary>
[DisplayName("Slide Transitions")]
[Description("Sets a slide transition.")]
public sealed class SlideTransitions : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>Which transition effect the slide enters with.</summary>
    [Category("Input")]
    [DisplayName("Entry Effect")]
    [Description("Which transition effect the slide enters with.")]
    public EntryEffectEnum EntryEffect { get; set; } = EntryEffectEnum.None;

    /// <summary>Advance the slide on a mouse click.</summary>
    [Category("Input")]
    [DisplayName("On Mouse Click")]
    [Description("Advance the slide on a mouse click.")]
    public bool OnMouseClick { get; set; }

    /// <summary>How long the transition runs, in seconds. Zero leaves the default.</summary>
    [Category("Input")]
    [DisplayName("Duration")]
    [Description("How long the transition runs, in seconds. Zero leaves the default.")]
    public InArgument<float> Duration { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.SlideTransitions(
            SlideIndex.Get(context),
            EntryEffect,
            OnMouseClick,
            Duration?.Get(context) ?? 0);
}
