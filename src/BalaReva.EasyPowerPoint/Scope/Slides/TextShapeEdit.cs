using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.PowerPoint;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Restyles one text shape on a slide.</summary>
/// <remarks>
/// TextStyle is carried on the activity because the published package declared
/// it there, but the preset it names maps to no single font setting, so the
/// service is handed an empty style. See docs/REVIVAL.md.
/// </remarks>
[DisplayName("Text Shape Edit")]
[Description("Restyles one text shape on a slide.")]
public sealed class TextShapeEdit : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>Which text shape on the slide. Text shapes are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Text Index")]
    [Description("Which text shape on the slide. Text shapes are numbered from 1.")]
    public InArgument<int> TextIndex { get; set; } = null!;

    /// <summary>Which preset style to apply to the shape text.</summary>
    [Category("Input")]
    [DisplayName("Text Style")]
    [Description("Which preset style to apply to the shape text.")]
    public ShapColorTypeEnum TextStyle { get; set; } = ShapColorTypeEnum.Mixed;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.TextShapeEdit(
            SlideIndex.Get(context),
            TextIndex.Get(context),
            new TextStyleRequest());
}
