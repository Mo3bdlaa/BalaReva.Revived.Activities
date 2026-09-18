using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.PowerPoint;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Restyles one text shape on a slide.</summary>
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

    /// <summary>The text, font and position to give the shape.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Text Style")]
    [Description("The text, font and position to give the shape.")]
    public InArgument<TextShape> TextStyle { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.TextShapeEdit(
            SlideIndex.Get(context),
            TextIndex.Get(context),
            TextStyle?.Get(context)
                ?? throw new ArgumentException("Text Style is required.", nameof(TextStyle)));
}
