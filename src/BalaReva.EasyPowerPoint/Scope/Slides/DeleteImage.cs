using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Deletes an image from a slide.</summary>
[DisplayName("Delete Image")]
[Description("Deletes an image from a slide.")]
public sealed class DeleteImage : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>Which image on the slide. Images are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Image Index")]
    [Description("Which image on the slide. Images are numbered from 1.")]
    public InArgument<int> ImageIndex { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.DeleteImage(SlideIndex.Get(context), ImageIndex.Get(context));
}
